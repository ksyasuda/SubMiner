import type { ConfigHotReloadDiff } from '../../core/services/config-hot-reload';
import { resolveKeybindings } from '../../core/utils/keybindings';
import { DEFAULT_KEYBINDINGS } from '../../config';
import type {
  ConfigHotReloadPayload,
  OverlayNotificationPayload,
  ResolvedConfig,
  SecondarySubMode,
} from '../../types';

type ConfigHotReloadAppliedDeps = {
  setKeybindings: (keybindings: ConfigHotReloadPayload['keybindings']) => void;
  refreshGlobalAndOverlayShortcuts: () => void;
  setSecondarySubMode: (mode: SecondarySubMode) => void;
  broadcastToOverlayWindows: (channel: string, payload: unknown) => void;
  applyAnkiRuntimeConfigPatch: (patch: {
    ai: ResolvedConfig['ankiConnect']['ai']['enabled'];
  }) => void;
};

type ConfigHotReloadMessageDeps = {
  showConfiguredNotification: (title: string, payload: OverlayNotificationPayload) => void;
};

export function resolveSubtitleStyleForRenderer(config: ResolvedConfig) {
  if (!config.subtitleStyle) {
    return null;
  }
  return {
    ...config.subtitleStyle,
    nPlusOneColor: config.ankiConnect.nPlusOne.nPlusOne,
    knownWordColor: config.ankiConnect.nPlusOne.knownWord,
    nameMatchColor: config.subtitleStyle.nameMatchColor,
    enableJlpt: config.subtitleStyle.enableJlpt,
    frequencyDictionary: config.subtitleStyle.frequencyDictionary,
  };
}

export function buildConfigHotReloadPayload(config: ResolvedConfig): ConfigHotReloadPayload {
  return {
    keybindings: resolveKeybindings(config, DEFAULT_KEYBINDINGS),
    subtitleStyle: resolveSubtitleStyleForRenderer(config),
    secondarySubMode: config.secondarySub.defaultMode,
  };
}

export function createConfigHotReloadAppliedHandler(deps: ConfigHotReloadAppliedDeps) {
  return (diff: ConfigHotReloadDiff, config: ResolvedConfig): void => {
    const payload = buildConfigHotReloadPayload(config);
    deps.setKeybindings(payload.keybindings);

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
    deps.showConfiguredNotification('SubMiner', {
      kind: 'warning',
      message,
    });
  };
}

export function buildRestartRequiredConfigMessage(fields: string[]): string {
  return `Config updated; restart required for: ${fields.join(', ')}`;
}
