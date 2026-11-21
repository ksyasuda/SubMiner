import type { SessionActionId, SessionActionPayload } from '../plugin-api';
export type { SessionActionId, SessionActionPayload } from '../plugin-api';

export type SessionKeyModifier = 'ctrl' | 'alt' | 'shift' | 'meta';

export interface SessionKeySpec {
  code: string;
  modifiers: SessionKeyModifier[];
}

export interface MpvInputBindingsSnapshot {
  keys: string[];
  blockedKeys: SessionKeySpec[];
}

/** Why the overlay kept a key that mpv binds instead of forwarding it. */
export type MpvKeyForwardBlock =
  | 'config-loading'
  | 'modal'
  | 'lookup-popup'
  | 'interactive-target'
  | 'no-imported-bindings';

/** Debug report the overlay sends when it does not forward an mpv-bound key. */
export interface UnforwardedMpvKeyReport {
  key: string;
  reason: MpvKeyForwardBlock;
}

export interface SessionBindingWarning {
  kind: 'unsupported' | 'conflict' | 'deprecated-config';
  path: string;
  message: string;
  value: unknown;
  conflictingPaths?: string[];
}

type CompiledSessionBindingBase = {
  sourcePath: string;
  originalKey: string;
  key: SessionKeySpec;
};

export interface CompiledMpvCommandBinding extends CompiledSessionBindingBase {
  actionType: 'mpv-command';
  command: (string | number)[];
}

export interface CompiledSessionActionBinding extends CompiledSessionBindingBase {
  actionType: 'session-action';
  actionId: SessionActionId;
  payload?: SessionActionPayload;
}

export type CompiledSessionBinding = CompiledMpvCommandBinding | CompiledSessionActionBinding;

export interface PluginSessionActionBinding extends CompiledSessionActionBinding {
  cliArgs?: string[];
}

export type PluginSessionBinding = CompiledMpvCommandBinding | PluginSessionActionBinding;

export interface PluginSessionBindingsArtifact {
  version: 1;
  generatedAt: string;
  numericSelectionTimeoutMs: number;
  bindings: PluginSessionBinding[];
  warnings: SessionBindingWarning[];
}
