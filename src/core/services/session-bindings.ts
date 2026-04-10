import type { Keybinding, ResolvedConfig } from '../../types';
import type { ConfiguredShortcuts } from '../utils/shortcut-config';
import type {
  CompiledMpvCommandBinding,
  CompiledSessionActionBinding,
  CompiledSessionBinding,
  PluginSessionBindingsArtifact,
  SessionActionId,
  SessionBindingWarning,
  SessionKeyModifier,
  SessionKeySpec,
} from '../../types/session-bindings';
import { SPECIAL_COMMANDS } from '../../config';

type PlatformKeyModel = 'darwin' | 'win32' | 'linux';

type CompileSessionBindingsInput = {
  keybindings: Keybinding[];
  shortcuts: ConfiguredShortcuts;
  statsToggleKey?: string | null;
  platform: PlatformKeyModel;
  rawConfig?: ResolvedConfig | null;
};

type DraftBinding = {
  binding: CompiledSessionBinding;
  actionFingerprint: string;
};

const MODIFIER_ORDER: SessionKeyModifier[] = ['ctrl', 'alt', 'shift', 'meta'];

const SESSION_SHORTCUT_ACTIONS: Array<{
  key: keyof Omit<ConfiguredShortcuts, 'multiCopyTimeoutMs'>;
  actionId: SessionActionId;
}> = [
  { key: 'toggleVisibleOverlayGlobal', actionId: 'toggleVisibleOverlay' },
  { key: 'copySubtitle', actionId: 'copySubtitle' },
  { key: 'copySubtitleMultiple', actionId: 'copySubtitleMultiple' },
  { key: 'updateLastCardFromClipboard', actionId: 'updateLastCardFromClipboard' },
  { key: 'triggerFieldGrouping', actionId: 'triggerFieldGrouping' },
  { key: 'triggerSubsync', actionId: 'triggerSubsync' },
  { key: 'mineSentence', actionId: 'mineSentence' },
  { key: 'mineSentenceMultiple', actionId: 'mineSentenceMultiple' },
  { key: 'toggleSecondarySub', actionId: 'toggleSecondarySub' },
  { key: 'markAudioCard', actionId: 'markAudioCard' },
  { key: 'openRuntimeOptions', actionId: 'openRuntimeOptions' },
  { key: 'openJimaku', actionId: 'openJimaku' },
];

function normalizeModifiers(modifiers: SessionKeyModifier[]): SessionKeyModifier[] {
  return [...new Set(modifiers)].sort(
    (left, right) => MODIFIER_ORDER.indexOf(left) - MODIFIER_ORDER.indexOf(right),
  );
}

function normalizeCodeToken(token: string): string | null {
  const normalized = token.trim();
  if (!normalized) return null;
  if (/^[a-z]$/i.test(normalized)) {
    return `Key${normalized.toUpperCase()}`;
  }
  if (/^[0-9]$/.test(normalized)) {
    return `Digit${normalized}`;
  }

  const exactMap: Record<string, string> = {
    space: 'Space',
    tab: 'Tab',
    enter: 'Enter',
    return: 'Enter',
    esc: 'Escape',
    escape: 'Escape',
    up: 'ArrowUp',
    down: 'ArrowDown',
    left: 'ArrowLeft',
    right: 'ArrowRight',
    backspace: 'Backspace',
    delete: 'Delete',
    slash: 'Slash',
    backslash: 'Backslash',
    minus: 'Minus',
    plus: 'Equal',
    equal: 'Equal',
    comma: 'Comma',
    period: 'Period',
    quote: 'Quote',
    semicolon: 'Semicolon',
    bracketleft: 'BracketLeft',
    bracketright: 'BracketRight',
    backquote: 'Backquote',
  };
  const lower = normalized.toLowerCase();
  if (exactMap[lower]) return exactMap[lower];
  if (
    /^key[a-z]$/i.test(normalized) ||
    /^digit[0-9]$/i.test(normalized) ||
    /^arrow(?:up|down|left|right)$/i.test(normalized) ||
    /^f\d{1,2}$/i.test(normalized)
  ) {
    return normalized[0]!.toUpperCase() + normalized.slice(1);
  }
  return null;
}

function parseAccelerator(
  accelerator: string,
  platform: PlatformKeyModel,
): { key: SessionKeySpec | null; message?: string } {
  const normalized = accelerator.replace(/\s+/g, '').replace(/cmdorctrl/gi, 'CommandOrControl');
  if (!normalized) {
    return { key: null, message: 'Empty accelerator is not supported.' };
  }

  const parts = normalized.split('+').filter(Boolean);
  const keyToken = parts.pop();
  if (!keyToken) {
    return { key: null, message: 'Missing accelerator key token.' };
  }

  const modifiers: SessionKeyModifier[] = [];
  for (const modifier of parts) {
    const lower = modifier.toLowerCase();
    if (lower === 'ctrl' || lower === 'control') {
      modifiers.push('ctrl');
      continue;
    }
    if (lower === 'alt' || lower === 'option') {
      modifiers.push('alt');
      continue;
    }
    if (lower === 'shift') {
      modifiers.push('shift');
      continue;
    }
    if (lower === 'meta' || lower === 'super' || lower === 'command' || lower === 'cmd') {
      modifiers.push('meta');
      continue;
    }
    if (lower === 'commandorcontrol') {
      modifiers.push(platform === 'darwin' ? 'meta' : 'ctrl');
      continue;
    }
    return {
      key: null,
      message: `Unsupported accelerator modifier: ${modifier}`,
    };
  }

  const code = normalizeCodeToken(keyToken);
  if (!code) {
    return {
      key: null,
      message: `Unsupported accelerator key token: ${keyToken}`,
    };
  }

  return {
    key: {
      code,
      modifiers: normalizeModifiers(modifiers),
    },
  };
}

function parseDomKeyString(
  key: string,
  platform: PlatformKeyModel,
): { key: SessionKeySpec | null; message?: string } {
  const parts = key
    .split('+')
    .map((part) => part.trim())
    .filter(Boolean);
  const keyToken = parts.pop();
  if (!keyToken) {
    return { key: null, message: 'Missing keybinding key token.' };
  }

  const modifiers: SessionKeyModifier[] = [];
  for (const modifier of parts) {
    const lower = modifier.toLowerCase();
    if (lower === 'ctrl' || lower === 'control') {
      modifiers.push('ctrl');
      continue;
    }
    if (lower === 'alt' || lower === 'option') {
      modifiers.push('alt');
      continue;
    }
    if (lower === 'shift') {
      modifiers.push('shift');
      continue;
    }
    if (
      lower === 'meta' ||
      lower === 'super' ||
      lower === 'command' ||
      lower === 'cmd' ||
      lower === 'commandorcontrol'
    ) {
      modifiers.push(
        lower === 'commandorcontrol' ? (platform === 'darwin' ? 'meta' : 'ctrl') : 'meta',
      );
      continue;
    }
    return {
      key: null,
      message: `Unsupported keybinding modifier: ${modifier}`,
    };
  }

  const code = normalizeCodeToken(keyToken);
  if (!code) {
    return {
      key: null,
      message: `Unsupported keybinding token: ${keyToken}`,
    };
  }

  return {
    key: {
      code,
      modifiers: normalizeModifiers(modifiers),
    },
  };
}

export function getSessionKeySpecSignature(key: SessionKeySpec): string {
  return [...key.modifiers, key.code].join('+');
}

function resolveCommandBinding(
  binding: Keybinding,
):
  | Omit<CompiledMpvCommandBinding, 'key' | 'sourcePath' | 'originalKey'>
  | Omit<CompiledSessionActionBinding, 'key' | 'sourcePath' | 'originalKey'>
  | null {
  const command = binding.command ?? [];
  const first = typeof command[0] === 'string' ? command[0] : '';
  if (first === SPECIAL_COMMANDS.SUBSYNC_TRIGGER) {
    return { actionType: 'session-action', actionId: 'triggerSubsync' };
  }
  if (first === SPECIAL_COMMANDS.RUNTIME_OPTIONS_OPEN) {
    return { actionType: 'session-action', actionId: 'openRuntimeOptions' };
  }
  if (first === SPECIAL_COMMANDS.JIMAKU_OPEN) {
    return { actionType: 'session-action', actionId: 'openJimaku' };
  }
  if (first === SPECIAL_COMMANDS.YOUTUBE_PICKER_OPEN) {
    return { actionType: 'session-action', actionId: 'openYoutubePicker' };
  }
  if (first === SPECIAL_COMMANDS.PLAYLIST_BROWSER_OPEN) {
    return { actionType: 'session-action', actionId: 'openPlaylistBrowser' };
  }
  if (first === SPECIAL_COMMANDS.REPLAY_SUBTITLE) {
    return { actionType: 'session-action', actionId: 'replayCurrentSubtitle' };
  }
  if (first === SPECIAL_COMMANDS.PLAY_NEXT_SUBTITLE) {
    return { actionType: 'session-action', actionId: 'playNextSubtitle' };
  }
  if (first === SPECIAL_COMMANDS.SHIFT_SUB_DELAY_TO_PREVIOUS_SUBTITLE_START) {
    return { actionType: 'session-action', actionId: 'shiftSubDelayPrevLine' };
  }
  if (first === SPECIAL_COMMANDS.SHIFT_SUB_DELAY_TO_NEXT_SUBTITLE_START) {
    return { actionType: 'session-action', actionId: 'shiftSubDelayNextLine' };
  }
  if (first.startsWith(SPECIAL_COMMANDS.RUNTIME_OPTION_CYCLE_PREFIX)) {
    const [, runtimeOptionId, rawDirection] = first.split(':');
    return {
      actionType: 'session-action',
      actionId: 'cycleRuntimeOption',
      payload: {
        runtimeOptionId,
        direction: rawDirection === 'prev' ? -1 : 1,
      },
    };
  }

  return {
    actionType: 'mpv-command',
    command,
  };
}

function getBindingFingerprint(binding: CompiledSessionBinding): string {
  if (binding.actionType === 'mpv-command') {
    return `mpv:${JSON.stringify(binding.command)}`;
  }
  return `session:${binding.actionId}:${JSON.stringify(binding.payload ?? null)}`;
}

export function compileSessionBindings(
  input: CompileSessionBindingsInput,
): {
  bindings: CompiledSessionBinding[];
  warnings: SessionBindingWarning[];
} {
  const warnings: SessionBindingWarning[] = [];
  const candidates = new Map<string, DraftBinding[]>();
  const legacyToggleVisibleOverlayGlobal = (
    input.rawConfig?.shortcuts as Record<string, unknown> | undefined
  )?.toggleVisibleOverlayGlobal;
  const statsToggleKey = input.statsToggleKey ?? input.rawConfig?.stats?.toggleKey ?? null;

  if (legacyToggleVisibleOverlayGlobal !== undefined) {
    warnings.push({
      kind: 'deprecated-config',
      path: 'shortcuts.toggleVisibleOverlayGlobal',
      value: legacyToggleVisibleOverlayGlobal,
      message: 'Rename shortcuts.toggleVisibleOverlayGlobal to shortcuts.toggleVisibleOverlay.',
    });
  }

  for (const shortcut of SESSION_SHORTCUT_ACTIONS) {
    const accelerator = input.shortcuts[shortcut.key];
    if (!accelerator) continue;
    const parsed = parseAccelerator(accelerator, input.platform);
    if (!parsed.key) {
      warnings.push({
        kind: 'unsupported',
        path: `shortcuts.${shortcut.key}`,
        value: accelerator,
        message: parsed.message ?? 'Unsupported accelerator syntax.',
      });
      continue;
    }
    const binding: CompiledSessionActionBinding = {
      sourcePath: `shortcuts.${shortcut.key}`,
      originalKey: accelerator,
      key: parsed.key,
      actionType: 'session-action',
      actionId: shortcut.actionId,
    };
    const signature = getSessionKeySpecSignature(parsed.key);
    const draft = candidates.get(signature) ?? [];
    draft.push({
      binding,
      actionFingerprint: getBindingFingerprint(binding),
    });
    candidates.set(signature, draft);
  }

  if (statsToggleKey) {
    const parsed = parseDomKeyString(statsToggleKey, input.platform);
    if (!parsed.key) {
      warnings.push({
        kind: 'unsupported',
        path: 'stats.toggleKey',
        value: statsToggleKey,
        message: parsed.message ?? 'Unsupported stats toggle key syntax.',
      });
    } else {
      const binding: CompiledSessionActionBinding = {
        sourcePath: 'stats.toggleKey',
        originalKey: statsToggleKey,
        key: parsed.key,
        actionType: 'session-action',
        actionId: 'toggleStatsOverlay',
      };
      const signature = getSessionKeySpecSignature(parsed.key);
      const draft = candidates.get(signature) ?? [];
      draft.push({
        binding,
        actionFingerprint: getBindingFingerprint(binding),
      });
      candidates.set(signature, draft);
    }
  }

  input.keybindings.forEach((binding, index) => {
    if (!binding.command) return;
    const parsed = parseDomKeyString(binding.key, input.platform);
    if (!parsed.key) {
      warnings.push({
        kind: 'unsupported',
        path: `keybindings[${index}].key`,
        value: binding.key,
        message: parsed.message ?? 'Unsupported keybinding syntax.',
      });
      return;
    }
    const resolved = resolveCommandBinding(binding);
    if (!resolved) return;
    const compiled: CompiledSessionBinding = {
      sourcePath: `keybindings[${index}].key`,
      originalKey: binding.key,
      key: parsed.key,
      ...resolved,
    };
    const signature = getSessionKeySpecSignature(parsed.key);
    const draft = candidates.get(signature) ?? [];
    draft.push({
      binding: compiled,
      actionFingerprint: getBindingFingerprint(compiled),
    });
    candidates.set(signature, draft);
  });

  const bindings: CompiledSessionBinding[] = [];
  for (const [signature, draftBindings] of candidates.entries()) {
    const uniqueFingerprints = new Set(draftBindings.map((entry) => entry.actionFingerprint));
    if (uniqueFingerprints.size > 1) {
      warnings.push({
        kind: 'conflict',
        path: draftBindings[0]!.binding.sourcePath,
        value: signature,
        conflictingPaths: draftBindings.map((entry) => entry.binding.sourcePath),
        message: `Conflicting session bindings compile to ${signature}; SubMiner will bind neither action.`,
      });
      continue;
    }
    bindings.push(draftBindings[0]!.binding);
  }

  bindings.sort((left, right) => left.sourcePath.localeCompare(right.sourcePath));
  return { bindings, warnings };
}

export function buildPluginSessionBindingsArtifact(input: {
  bindings: CompiledSessionBinding[];
  warnings: SessionBindingWarning[];
  numericSelectionTimeoutMs: number;
  now?: Date;
}): PluginSessionBindingsArtifact {
  return {
    version: 1,
    generatedAt: (input.now ?? new Date()).toISOString(),
    numericSelectionTimeoutMs: input.numericSelectionTimeoutMs,
    bindings: input.bindings,
    warnings: input.warnings,
  };
}
