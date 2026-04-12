import type { ConfigHotReloadDiff } from '../../core/services/config-hot-reload';
import { compileSessionBindings } from '../../core/services/session-bindings';
import { resolveKeybindings } from '../../core/utils/keybindings';
import { resolveConfiguredShortcuts } from '../../core/utils/shortcut-config';
import { DEFAULT_CONFIG, DEFAULT_KEYBINDINGS } from '../../config';
import type { ConfigHotReloadPayload, ResolvedConfig, SecondarySubMode } from '../../types';

type ConfigHotReloadAppliedDeps = {
  setKeybindings: (keybindings: ConfigHotReloadPayload['keybindings']) => void;
  setSessionBindings: (
    sessionBindings: ConfigHotReloadPayload['sessionBindings'],
    sessionBindingWarnings: ConfigHotReloadPayload['sessionBindingWarnings'],
  ) => void;
  refreshGlobalAndOverlayShortcuts: () => void;
  setSecondarySubMode: (mode: SecondarySubMode) => void;
  broadcastToOverlayWindows: (channel: string, payload: unknown) => void;
  applyAnkiRuntimeConfigPatch: (patch: {
    ai: ResolvedConfig['ankiConnect']['ai']['enabled'];
  }) => void;
};

type ConfigHotReloadMessageDeps = {
  showMpvOsd: (message: string) => void;
  showDesktopNotification: (title: string, options: { body: string }) => void;
};

export function resolveSubtitleStyleForRenderer(config: ResolvedConfig) {
  if (!config.subtitleStyle) {
    return null;
  }
  return {
    ...config.subtitleStyle,
    nPlusOneColor: config.ankiConnect.nPlusOne.nPlusOne,
    knownWordColor: config.ankiConnect.knownWords.color,
    nameMatchColor: config.subtitleStyle.nameMatchColor,
    enableJlpt: config.subtitleStyle.enableJlpt,
    frequencyDictionary: config.subtitleStyle.frequencyDictionary,
  };
}

export function buildConfigHotReloadPayload(config: ResolvedConfig): ConfigHotReloadPayload {
  const keybindings = resolveKeybindings(config, DEFAULT_KEYBINDINGS);
  const { bindings: sessionBindings, warnings: sessionBindingWarnings } = compileSessionBindings({
    keybindings,
    shortcuts: resolveConfiguredShortcuts(config, DEFAULT_CONFIG),
    statsToggleKey: config.stats.toggleKey,
    platform:
      process.platform === 'darwin' ? 'darwin' : process.platform === 'win32' ? 'win32' : 'linux',
    rawConfig: config,
  });
  return {
    keybindings,
    sessionBindings,
    sessionBindingWarnings,
    subtitleStyle: resolveSubtitleStyleForRenderer(config),
    subtitleSidebar: config.subtitleSidebar,
    secondarySubMode: config.secondarySub.defaultMode,
  };
}

export function createConfigHotReloadAppliedHandler(deps: ConfigHotReloadAppliedDeps) {
  return (diff: ConfigHotReloadDiff, config: ResolvedConfig): void => {
    const payload = buildConfigHotReloadPayload(config);
    deps.setKeybindings(payload.keybindings);
    deps.setSessionBindings(payload.sessionBindings, payload.sessionBindingWarnings);

    if (diff.hotReloadFields.includes('shortcuts')) {
      deps.refreshGlobalAndOverlayShortcuts();
    }

    if (diff.hotReloadFields.includes('secondarySub.defaultMode')) {
      deps.setSecondarySubMode(payload.secondarySubMode);
      deps.broadcastToOverlayWindows('secondary-subtitle:mode', payload.secondarySubMode);
    }

    if (diff.hotReloadFields.includes('ankiConnect.ai')) {
      deps.applyAnkiRuntimeConfigPatch({ ai: config.ankiConnect.ai.enabled });
    }

    if (diff.hotReloadFields.length > 0) {
      deps.broadcastToOverlayWindows('config:hot-reload', payload);
    }
  };
}

export function createConfigHotReloadMessageHandler(deps: ConfigHotReloadMessageDeps) {
  return (message: string): void => {
    deps.showMpvOsd(message);
    deps.showDesktopNotification('SubMiner', { body: message });
  };
}

export function buildRestartRequiredConfigMessage(fields: string[]): string {
  return `Config updated; restart required for: ${fields.join(', ')}`;
}
