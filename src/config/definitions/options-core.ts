import { ResolvedConfig } from '../../types/config';
import { ConfigOptionRegistryEntry } from './shared';

export function buildCoreConfigOptionRegistry(
  defaultConfig: ResolvedConfig,
): ConfigOptionRegistryEntry[] {
  const discreteBindings = [
    {
      id: 'toggleLookup',
      defaultValue: defaultConfig.controller.bindings.toggleLookup,
      description: 'Controller binding descriptor for toggling lookup.',
    },
    {
      id: 'closeLookup',
      defaultValue: defaultConfig.controller.bindings.closeLookup,
      description: 'Controller binding descriptor for closing lookup.',
    },
    {
      id: 'toggleKeyboardOnlyMode',
      defaultValue: defaultConfig.controller.bindings.toggleKeyboardOnlyMode,
      description: 'Controller binding descriptor for toggling keyboard-only mode.',
    },
    {
      id: 'mineCard',
      defaultValue: defaultConfig.controller.bindings.mineCard,
      description: 'Controller binding descriptor for mining the active card.',
    },
    {
      id: 'quitMpv',
      defaultValue: defaultConfig.controller.bindings.quitMpv,
      description: 'Controller binding descriptor for quitting mpv.',
    },
    {
      id: 'previousAudio',
      defaultValue: defaultConfig.controller.bindings.previousAudio,
      description: 'Controller binding descriptor for previous Yomitan audio.',
    },
    {
      id: 'nextAudio',
      defaultValue: defaultConfig.controller.bindings.nextAudio,
      description: 'Controller binding descriptor for next Yomitan audio.',
    },
    {
      id: 'playCurrentAudio',
      defaultValue: defaultConfig.controller.bindings.playCurrentAudio,
      description: 'Controller binding descriptor for playing the current Yomitan audio.',
    },
    {
      id: 'toggleMpvPause',
      defaultValue: defaultConfig.controller.bindings.toggleMpvPause,
      description: 'Controller binding descriptor for toggling mpv play/pause.',
    },
  ] as const;

  const axisBindings = [
    {
      id: 'leftStickHorizontal',
      defaultValue: defaultConfig.controller.bindings.leftStickHorizontal,
      description: 'Axis binding descriptor used for left/right token selection.',
    },
    {
      id: 'leftStickVertical',
      defaultValue: defaultConfig.controller.bindings.leftStickVertical,
      description: 'Axis binding descriptor used for primary popup scrolling.',
    },
    {
      id: 'rightStickHorizontal',
      defaultValue: defaultConfig.controller.bindings.rightStickHorizontal,
      description: 'Axis binding descriptor reserved for alternate right-stick mappings.',
    },
    {
      id: 'rightStickVertical',
      defaultValue: defaultConfig.controller.bindings.rightStickVertical,
      description: 'Axis binding descriptor used for popup page jumps.',
    },
  ] as const;

  return [
    {
      path: 'logging.level',
      kind: 'enum',
      enumValues: ['debug', 'info', 'warn', 'error'],
      defaultValue: defaultConfig.logging.level,
      description: 'Minimum log level for runtime logging.',
    },
    {
      path: 'youtube.primarySubLanguages',
      kind: 'string',
      defaultValue: defaultConfig.youtube.primarySubLanguages.join(','),
      description:
        'Comma-separated primary subtitle language priority for managed subtitle auto-selection.',
    },
    {
      path: 'controller.enabled',
      kind: 'boolean',
      defaultValue: defaultConfig.controller.enabled,
      description: 'Enable overlay controller support through the Chrome Gamepad API.',
    },
    {
      path: 'controller.preferredGamepadId',
      kind: 'string',
      defaultValue: defaultConfig.controller.preferredGamepadId,
      description: 'Preferred controller id saved from the controller config modal.',
    },
    {
      path: 'controller.preferredGamepadLabel',
      kind: 'string',
      defaultValue: defaultConfig.controller.preferredGamepadLabel,
      description: 'Preferred controller display label saved for diagnostics.',
    },
    {
      path: 'controller.smoothScroll',
      kind: 'boolean',
      defaultValue: defaultConfig.controller.smoothScroll,
      description: 'Use smooth scrolling for controller-driven popup scroll input.',
    },
    {
      path: 'controller.scrollPixelsPerSecond',
      kind: 'number',
      defaultValue: defaultConfig.controller.scrollPixelsPerSecond,
      description: 'Base popup scroll speed for controller stick input.',
    },
    {
      path: 'controller.horizontalJumpPixels',
      kind: 'number',
      defaultValue: defaultConfig.controller.horizontalJumpPixels,
      description: 'Popup page-jump distance for controller jump input.',
    },
    {
      path: 'controller.stickDeadzone',
      kind: 'number',
      defaultValue: defaultConfig.controller.stickDeadzone,
      description: 'Deadzone applied to controller stick axes.',
    },
    {
      path: 'controller.triggerInputMode',
      kind: 'enum',
      enumValues: ['auto', 'digital', 'analog'],
      defaultValue: defaultConfig.controller.triggerInputMode,
      description:
        'How controller triggers are interpreted: auto, pressed-only, or thresholded analog.',
    },
    {
      path: 'controller.triggerDeadzone',
      kind: 'number',
      defaultValue: defaultConfig.controller.triggerDeadzone,
      description:
        'Minimum analog trigger value required when trigger input uses auto or analog mode.',
    },
    {
      path: 'controller.repeatDelayMs',
      kind: 'number',
      defaultValue: defaultConfig.controller.repeatDelayMs,
      description: 'Delay before repeating held controller actions.',
    },
    {
      path: 'controller.repeatIntervalMs',
      kind: 'number',
      defaultValue: defaultConfig.controller.repeatIntervalMs,
      description: 'Repeat interval for held controller actions.',
    },
    {
      path: 'controller.buttonIndices',
      kind: 'object',
      defaultValue: defaultConfig.controller.buttonIndices,
      description:
        'Semantic button-name reference mapping used for legacy configs and debug output. Updating it does not rewrite existing raw binding descriptors.',
    },
    {
      path: 'controller.buttonIndices.select',
      kind: 'number',
      defaultValue: defaultConfig.controller.buttonIndices.select,
      description: 'Raw button index used for the controller select/minus/back button.',
    },
    {
      path: 'controller.buttonIndices.buttonSouth',
      kind: 'number',
      defaultValue: defaultConfig.controller.buttonIndices.buttonSouth,
      description: 'Raw button index used for controller south/A button input.',
    },
    {
      path: 'controller.buttonIndices.buttonEast',
      kind: 'number',
      defaultValue: defaultConfig.controller.buttonIndices.buttonEast,
      description: 'Raw button index used for controller east/B button input.',
    },
    {
      path: 'controller.buttonIndices.buttonNorth',
      kind: 'number',
      defaultValue: defaultConfig.controller.buttonIndices.buttonNorth,
      description: 'Raw button index used for controller north/Y button input.',
    },
    {
      path: 'controller.buttonIndices.buttonWest',
      kind: 'number',
      defaultValue: defaultConfig.controller.buttonIndices.buttonWest,
      description: 'Raw button index used for controller west/X button input.',
    },
    {
      path: 'controller.buttonIndices.leftShoulder',
      kind: 'number',
      defaultValue: defaultConfig.controller.buttonIndices.leftShoulder,
      description: 'Raw button index used for controller left shoulder input.',
    },
    {
      path: 'controller.buttonIndices.rightShoulder',
      kind: 'number',
      defaultValue: defaultConfig.controller.buttonIndices.rightShoulder,
      description: 'Raw button index used for controller right shoulder input.',
    },
    {
      path: 'controller.buttonIndices.leftStickPress',
      kind: 'number',
      defaultValue: defaultConfig.controller.buttonIndices.leftStickPress,
      description: 'Raw button index used for controller L3 input.',
    },
    {
      path: 'controller.buttonIndices.rightStickPress',
      kind: 'number',
      defaultValue: defaultConfig.controller.buttonIndices.rightStickPress,
      description: 'Raw button index used for controller R3 input.',
    },
    {
      path: 'controller.buttonIndices.leftTrigger',
      kind: 'number',
      defaultValue: defaultConfig.controller.buttonIndices.leftTrigger,
      description: 'Raw button index used for controller L2 input.',
    },
    {
      path: 'controller.buttonIndices.rightTrigger',
      kind: 'number',
      defaultValue: defaultConfig.controller.buttonIndices.rightTrigger,
      description: 'Raw button index used for controller R2 input.',
    },
    {
      path: 'controller.bindings',
      kind: 'object',
      defaultValue: defaultConfig.controller.bindings,
      description:
        'Raw controller binding descriptors saved by Alt+C learn mode. For discrete axis bindings, kind "axis" requires axisIndex and direction.',
    },
    {
      path: 'controller.profiles',
      kind: 'object',
      defaultValue: defaultConfig.controller.profiles,
      description:
        'Per-controller binding and button-index overrides keyed by the controller id reported by the Gamepad API.',
    },
    ...discreteBindings.flatMap((binding) => [
      {
        path: `controller.bindings.${binding.id}`,
        kind: 'object' as const,
        defaultValue: binding.defaultValue,
        description: `${binding.description} Use Alt+C learn mode or set a raw button/axis descriptor manually. If kind is "axis", direction is required.`,
      },
      {
        path: `controller.bindings.${binding.id}.kind`,
        kind: 'enum' as const,
        enumValues: ['none', 'button', 'axis'],
        defaultValue: binding.defaultValue.kind,
        description:
          'Discrete binding input source kind. When kind is "axis", set both axisIndex and direction.',
      },
      {
        path: `controller.bindings.${binding.id}.buttonIndex`,
        kind: 'number' as const,
        defaultValue:
          binding.defaultValue.kind === 'button' ? binding.defaultValue.buttonIndex : undefined,
        description: 'Raw button index captured for this discrete controller action.',
      },
      {
        path: `controller.bindings.${binding.id}.axisIndex`,
        kind: 'number' as const,
        defaultValue:
          binding.defaultValue.kind === 'axis' ? binding.defaultValue.axisIndex : undefined,
        description: 'Raw axis index captured for this discrete controller action.',
      },
      {
        path: `controller.bindings.${binding.id}.direction`,
        kind: 'enum' as const,
        enumValues: ['negative', 'positive'],
        defaultValue:
          binding.defaultValue.kind === 'axis' ? binding.defaultValue.direction : undefined,
        description:
          'Axis direction captured for this discrete controller action. Required when kind is "axis".',
      },
    ]),
    ...axisBindings.flatMap((binding) => [
      {
        path: `controller.bindings.${binding.id}`,
        kind: 'object' as const,
        defaultValue: binding.defaultValue,
        description: `${binding.description} Use Alt+C learn mode or set a raw axis descriptor manually.`,
      },
      {
        path: `controller.bindings.${binding.id}.kind`,
        kind: 'enum' as const,
        enumValues: ['none', 'axis'],
        defaultValue: binding.defaultValue.kind,
        description: 'Analog binding input source kind.',
      },
      {
        path: `controller.bindings.${binding.id}.axisIndex`,
        kind: 'number' as const,
        defaultValue:
          binding.defaultValue.kind === 'axis' ? binding.defaultValue.axisIndex : undefined,
        description: 'Raw axis index captured for this analog controller action.',
      },
      {
        path: `controller.bindings.${binding.id}.dpadFallback`,
        kind: 'enum' as const,
        enumValues: ['none', 'horizontal', 'vertical'],
        defaultValue:
          binding.defaultValue.kind === 'axis' ? binding.defaultValue.dpadFallback : undefined,
        description:
          'Optional D-pad fallback used when this analog controller action should also read D-pad input.',
      },
    ]),
    {
      path: 'texthooker.launchAtStartup',
      kind: 'boolean',
      defaultValue: defaultConfig.texthooker.launchAtStartup,
      description: 'Launch texthooker server automatically when SubMiner starts.',
    },
    {
      path: 'texthooker.openBrowser',
      kind: 'boolean',
      defaultValue: defaultConfig.texthooker.openBrowser,
      description: 'Open the texthooker page in the default browser when the server starts.',
    },
    {
      path: 'subtitlePosition.yPercent',
      kind: 'number',
      defaultValue: defaultConfig.subtitlePosition.yPercent,
      description:
        'Vertical position of the subtitle overlay expressed as a percentage from the bottom of the screen.',
    },
    {
      path: 'auto_start_overlay',
      kind: 'boolean',
      defaultValue: defaultConfig.auto_start_overlay,
      description:
        'Show the visible subtitle overlay automatically when the bundled mpv plugin starts SubMiner.',
    },
    {
      path: 'secondarySub.secondarySubLanguages',
      kind: 'array',
      defaultValue: defaultConfig.secondarySub.secondarySubLanguages,
      description:
        'Language code priority list used to auto-select a secondary subtitle track when available.',
    },
    {
      path: 'secondarySub.autoLoadSecondarySub',
      kind: 'boolean',
      defaultValue: defaultConfig.secondarySub.autoLoadSecondarySub,
      description:
        'Automatically load a matching secondary subtitle when the primary subtitle loads.',
    },
    {
      path: 'secondarySub.defaultMode',
      kind: 'enum',
      enumValues: ['hidden', 'visible', 'hover'],
      defaultValue: defaultConfig.secondarySub.defaultMode,
      description: 'Default visibility mode for the secondary subtitle bar.',
    },
    {
      path: 'websocket.enabled',
      kind: 'enum',
      enumValues: ['auto', 'true', 'false'],
      defaultValue: defaultConfig.websocket.enabled,
      description: 'Built-in subtitle websocket server mode.',
    },
    {
      path: 'websocket.port',
      kind: 'number',
      defaultValue: defaultConfig.websocket.port,
      description: 'Built-in subtitle websocket server port.',
    },
    {
      path: 'annotationWebsocket.enabled',
      kind: 'boolean',
      defaultValue: defaultConfig.annotationWebsocket.enabled,
      description: 'Annotated subtitle websocket server enabled state.',
    },
    {
      path: 'annotationWebsocket.port',
      kind: 'number',
      defaultValue: defaultConfig.annotationWebsocket.port,
      description: 'Annotated subtitle websocket server port.',
    },
    {
      path: 'subsync.replace',
      kind: 'boolean',
      defaultValue: defaultConfig.subsync.replace,
      description: 'Replace the active subtitle file when sync completes.',
    },
    {
      path: 'subsync.alass_path',
      kind: 'string',
      defaultValue: defaultConfig.subsync.alass_path,
      description:
        'Optional absolute path to the alass binary used by subsync. Leave empty to auto-discover from PATH.',
    },
    {
      path: 'subsync.ffsubsync_path',
      kind: 'string',
      defaultValue: defaultConfig.subsync.ffsubsync_path,
      description:
        'Optional absolute path to the ffsubsync binary used by subsync. Leave empty to auto-discover from PATH.',
    },
    {
      path: 'subsync.ffmpeg_path',
      kind: 'string',
      defaultValue: defaultConfig.subsync.ffmpeg_path,
      description:
        'Optional absolute path to the ffmpeg binary used by subsync. Leave empty to auto-discover from PATH.',
    },
    {
      path: 'startupWarmups.lowPowerMode',
      kind: 'boolean',
      defaultValue: defaultConfig.startupWarmups.lowPowerMode,
      description: 'Defer startup warmups except Yomitan extension.',
    },
    {
      path: 'startupWarmups.mecab',
      kind: 'boolean',
      defaultValue: defaultConfig.startupWarmups.mecab,
      description: 'Warm up MeCab tokenizer at startup.',
    },
    {
      path: 'startupWarmups.yomitanExtension',
      kind: 'boolean',
      defaultValue: defaultConfig.startupWarmups.yomitanExtension,
      description: 'Warm up Yomitan extension at startup.',
    },
    {
      path: 'startupWarmups.subtitleDictionaries',
      kind: 'boolean',
      defaultValue: defaultConfig.startupWarmups.subtitleDictionaries,
      description: 'Warm up subtitle dictionaries at startup.',
    },
    {
      path: 'startupWarmups.jellyfinRemoteSession',
      kind: 'boolean',
      defaultValue: defaultConfig.startupWarmups.jellyfinRemoteSession,
      description: 'Warm up Jellyfin remote session at startup.',
    },
    {
      path: 'updates.enabled',
      kind: 'boolean',
      defaultValue: defaultConfig.updates.enabled,
      description: 'Run automatic update checks in the background.',
    },
    {
      path: 'updates.checkIntervalHours',
      kind: 'number',
      defaultValue: defaultConfig.updates.checkIntervalHours,
      description: 'Minimum hours between automatic update checks.',
    },
    {
      path: 'updates.notificationType',
      kind: 'enum',
      enumValues: ['system', 'osd', 'both', 'none'],
      defaultValue: defaultConfig.updates.notificationType,
      description: 'How SubMiner announces available updates.',
    },
    {
      path: 'updates.channel',
      kind: 'enum',
      enumValues: ['stable', 'prerelease'],
      defaultValue: defaultConfig.updates.channel,
      description: 'Release channel used for update checks.',
    },
    {
      path: 'shortcuts.multiCopyTimeoutMs',
      kind: 'number',
      defaultValue: defaultConfig.shortcuts.multiCopyTimeoutMs,
      description: 'Timeout for multi-copy/mine modes.',
    },
    {
      path: 'shortcuts.toggleVisibleOverlayGlobal',
      kind: 'string',
      defaultValue: defaultConfig.shortcuts.toggleVisibleOverlayGlobal,
      description:
        'Global accelerator that toggles overlay visibility from anywhere on the system. Use null to disable.',
    },
    {
      path: 'shortcuts.copySubtitle',
      kind: 'string',
      defaultValue: defaultConfig.shortcuts.copySubtitle,
      description: 'Accelerator that copies the current subtitle line to the clipboard.',
    },
    {
      path: 'shortcuts.copySubtitleMultiple',
      kind: 'string',
      defaultValue: defaultConfig.shortcuts.copySubtitleMultiple,
      description:
        'Accelerator that copies consecutive subtitle lines while the multi-copy window stays open.',
    },
    {
      path: 'shortcuts.updateLastCardFromClipboard',
      kind: 'string',
      defaultValue: defaultConfig.shortcuts.updateLastCardFromClipboard,
      description:
        'Accelerator that updates the last mined Anki card using the current clipboard contents.',
    },
    {
      path: 'shortcuts.triggerFieldGrouping',
      kind: 'string',
      defaultValue: defaultConfig.shortcuts.triggerFieldGrouping,
      description: 'Accelerator that triggers Kiku field grouping on duplicate cards.',
    },
    {
      path: 'shortcuts.triggerSubsync',
      kind: 'string',
      defaultValue: defaultConfig.shortcuts.triggerSubsync,
      description: 'Accelerator that triggers subsync against the active subtitle file.',
    },
    {
      path: 'shortcuts.mineSentence',
      kind: 'string',
      defaultValue: defaultConfig.shortcuts.mineSentence,
      description: 'Accelerator that mines the current sentence as a new Anki card.',
    },
    {
      path: 'shortcuts.mineSentenceMultiple',
      kind: 'string',
      defaultValue: defaultConfig.shortcuts.mineSentenceMultiple,
      description:
        'Accelerator that mines consecutive sentences while the multi-mine window stays open.',
    },
    {
      path: 'shortcuts.toggleSecondarySub',
      kind: 'string',
      defaultValue: defaultConfig.shortcuts.toggleSecondarySub,
      description: 'Accelerator that toggles the secondary subtitle bar visibility.',
    },
    {
      path: 'shortcuts.markAudioCard',
      kind: 'string',
      defaultValue: defaultConfig.shortcuts.markAudioCard,
      description: 'Accelerator that marks the last mined card as an audio card.',
    },
    {
      path: 'shortcuts.openCharacterDictionaryManager',
      kind: 'string',
      defaultValue: defaultConfig.shortcuts.openCharacterDictionaryManager,
      description: 'Accelerator that opens the character dictionary manager modal.',
    },
    {
      path: 'shortcuts.openRuntimeOptions',
      kind: 'string',
      defaultValue: defaultConfig.shortcuts.openRuntimeOptions,
      description: 'Accelerator that opens the runtime options modal.',
    },
    {
      path: 'shortcuts.openJimaku',
      kind: 'string',
      defaultValue: defaultConfig.shortcuts.openJimaku,
      description: 'Accelerator that opens the Jimaku subtitle search modal.',
    },
    {
      path: 'shortcuts.openSessionHelp',
      kind: 'string',
      defaultValue: defaultConfig.shortcuts.openSessionHelp,
      description: 'Accelerator that opens the session help / keybinding cheatsheet.',
    },
    {
      path: 'shortcuts.openControllerSelect',
      kind: 'string',
      defaultValue: defaultConfig.shortcuts.openControllerSelect,
      description: 'Accelerator that opens the controller selection and learn-mode modal.',
    },
    {
      path: 'shortcuts.openControllerDebug',
      kind: 'string',
      defaultValue: defaultConfig.shortcuts.openControllerDebug,
      description:
        'Accelerator that opens the controller debug modal with live axis/button readouts.',
    },
    {
      path: 'shortcuts.toggleSubtitleSidebar',
      kind: 'string',
      defaultValue: defaultConfig.shortcuts.toggleSubtitleSidebar,
      description: 'Accelerator that toggles the subtitle sidebar visibility.',
    },
  ];
}
