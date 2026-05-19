import type { ConfigHotReloadDiff } from '../../core/services/config-hot-reload';
import { compileSessionBindings } from '../../core/services/session-bindings';
import { resolveKeybindings } from '../../core/utils/keybindings';
import { resolveConfiguredShortcuts } from '../../core/utils/shortcut-config';
import { DEFAULT_CONFIG, DEFAULT_KEYBINDINGS } from '../../config';
import type { AnkiConnectConfig } from '../../types/anki';
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
  applyAnkiRuntimeConfigPatch: (patch: Partial<AnkiConnectConfig>) => void;
  invalidateTokenizationCache?: () => void;
  refreshSubtitlePrefetch?: () => void;
  refreshCurrentSubtitle?: () => void;
  setLogLevel?: (level: ResolvedConfig['logging']['level']) => void;
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
    nPlusOneColor: config.subtitleStyle.nPlusOneColor,
    knownWordColor: config.subtitleStyle.knownWordColor,
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
    primarySubMode: config.subtitleStyle.primaryDefaultMode,
    secondarySubMode: config.secondarySub.defaultMode,
  };
}

function hasAnyHotReloadField(diff: ConfigHotReloadDiff, prefixes: string[]): boolean {
  return diff.hotReloadFields.some((field) =>
    prefixes.some((prefix) => field === prefix || field.startsWith(`${prefix}.`)),
  );
}

function buildAnkiRuntimeConfigPatch(
  diff: ConfigHotReloadDiff,
  config: ResolvedConfig,
): Partial<AnkiConnectConfig> | null {
  const patch: Partial<AnkiConnectConfig> = {};

  if (diff.hotReloadFields.includes('ankiConnect.ai')) {
    patch.ai = config.ankiConnect.ai.enabled;
  }
  if (diff.hotReloadFields.includes('ankiConnect.ai.enabled')) {
    patch.ai = config.ankiConnect.ai.enabled;
  }
  if (diff.hotReloadFields.includes('ankiConnect.behavior.autoUpdateNewCards')) {
    patch.behavior = { autoUpdateNewCards: config.ankiConnect.behavior.autoUpdateNewCards };
  }
  if (hasAnyHotReloadField(diff, ['ankiConnect.knownWords'])) {
    patch.knownWords = config.ankiConnect.knownWords;
  }
  if (hasAnyHotReloadField(diff, ['ankiConnect.nPlusOne'])) {
    patch.nPlusOne = config.ankiConnect.nPlusOne;
  }
  const fieldPatch: NonNullable<AnkiConnectConfig['fields']> = {};
  if (diff.hotReloadFields.includes('ankiConnect.fields.word')) {
    fieldPatch.word = config.ankiConnect.fields.word;
  }
  if (diff.hotReloadFields.includes('ankiConnect.fields.audio')) {
    fieldPatch.audio = config.ankiConnect.fields.audio;
  }
  if (diff.hotReloadFields.includes('ankiConnect.fields.image')) {
    fieldPatch.image = config.ankiConnect.fields.image;
  }
  if (diff.hotReloadFields.includes('ankiConnect.fields.sentence')) {
    fieldPatch.sentence = config.ankiConnect.fields.sentence;
  }
  if (diff.hotReloadFields.includes('ankiConnect.fields.miscInfo')) {
    fieldPatch.miscInfo = config.ankiConnect.fields.miscInfo;
  }
  if (Object.keys(fieldPatch).length > 0) {
    patch.fields = fieldPatch;
  }
  if (diff.hotReloadFields.includes('ankiConnect.isLapis.sentenceCardModel')) {
    patch.isLapis = { sentenceCardModel: config.ankiConnect.isLapis.sentenceCardModel };
  }
  if (diff.hotReloadFields.includes('ankiConnect.isKiku.fieldGrouping')) {
    patch.isKiku = { fieldGrouping: config.ankiConnect.isKiku.fieldGrouping };
  }

  return Object.keys(patch).length > 0 ? patch : null;
}

function hasAnnotationRuntimeHotReload(diff: ConfigHotReloadDiff): boolean {
  return hasAnyHotReloadField(diff, [
    'ankiConnect.knownWords',
    'ankiConnect.nPlusOne',
    'ankiConnect.fields.word',
  ]);
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

    const ankiPatch = buildAnkiRuntimeConfigPatch(diff, config);
    if (ankiPatch) {
      deps.applyAnkiRuntimeConfigPatch(ankiPatch);
    }

    if (hasAnnotationRuntimeHotReload(diff)) {
      deps.invalidateTokenizationCache?.();
      deps.refreshSubtitlePrefetch?.();
      deps.refreshCurrentSubtitle?.();
    }

    if (diff.hotReloadFields.includes('logging.level')) {
      deps.setLogLevel?.(config.logging.level);
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
