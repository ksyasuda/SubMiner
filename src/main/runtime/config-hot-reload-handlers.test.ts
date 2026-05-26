import test from 'node:test';
import assert from 'node:assert/strict';
import { DEFAULT_CONFIG, deepCloneConfig } from '../../config';
import {
  buildConfigHotReloadPayload,
  buildRestartRequiredConfigMessage,
  createConfigHotReloadAppliedHandler,
  createConfigHotReloadMessageHandler,
} from './config-hot-reload-handlers';

test('createConfigHotReloadAppliedHandler runs all hot-reload effects', () => {
  const config = deepCloneConfig(DEFAULT_CONFIG);
  const calls: string[] = [];
  const ankiPatches: unknown[] = [];
  const sessionBindingWarnings: string[][] = [];

  const applyHotReload = createConfigHotReloadAppliedHandler({
    setKeybindings: () => calls.push('set:keybindings'),
    setSessionBindings: (_sessionBindings, warnings) => {
      calls.push('set:session-bindings');
      sessionBindingWarnings.push(warnings.map((warning) => warning.message));
    },
    refreshGlobalAndOverlayShortcuts: () => calls.push('refresh:shortcuts'),
    setSecondarySubMode: (mode) => calls.push(`set:secondary:${mode}`),
    broadcastToOverlayWindows: (channel, payload) =>
      calls.push(`broadcast:${channel}:${typeof payload === 'string' ? payload : 'object'}`),
    applyAnkiRuntimeConfigPatch: (patch) => {
      ankiPatches.push(patch);
    },
  });

  applyHotReload(
    {
      hotReloadFields: [
        'shortcuts',
        'secondarySub.defaultMode',
        'ankiConnect.ai',
        'subtitleStyle.autoPauseVideoOnHover',
      ],
      restartRequiredFields: [],
    },
    config,
  );

  assert.ok(calls.includes('set:keybindings'));
  assert.ok(calls.includes('set:session-bindings'));
  assert.ok(calls.includes('refresh:shortcuts'));
  assert.ok(calls.includes(`set:secondary:${config.secondarySub.defaultMode}`));
  assert.ok(calls.some((entry) => entry.startsWith('broadcast:secondary-subtitle:mode:')));
  assert.ok(calls.includes('broadcast:config:hot-reload:object'));
  assert.deepEqual(ankiPatches, [{ ai: config.ankiConnect.ai.enabled }]);
  assert.equal(sessionBindingWarnings.length, 1);
  assert.ok(
    sessionBindingWarnings[0]?.some((message) =>
      message.includes('Rename shortcuts.toggleVisibleOverlayGlobal'),
    ),
  );
});

test('createConfigHotReloadAppliedHandler applies safe Anki, annotation, and logging changes', () => {
  const config = deepCloneConfig(DEFAULT_CONFIG);
  config.ankiConnect.behavior.autoUpdateNewCards = false;
  config.ankiConnect.knownWords.highlightEnabled = true;
  config.ankiConnect.knownWords.refreshMinutes = 90;
  config.ankiConnect.knownWords.decks = { Anime: ['Mining'] };
  config.ankiConnect.nPlusOne.enabled = true;
  config.ankiConnect.nPlusOne.minSentenceWords = 4;
  config.ankiConnect.fields.word = 'Expression';
  config.ankiConnect.fields.audio = 'SentenceAudioCustom';
  config.ankiConnect.fields.image = 'ScreenshotCustom';
  config.ankiConnect.fields.sentence = 'SentenceCustom';
  config.ankiConnect.fields.miscInfo = 'MiscInfoCustom';
  config.ankiConnect.isLapis.sentenceCardModel = 'Sentence Card Custom';
  config.ankiConnect.isKiku.fieldGrouping = 'manual';
  config.logging.level = 'debug';
  config.logging.rotation = 14;
  config.logging.files.mpv = true;
  const calls: string[] = [];
  const ankiPatches: unknown[] = [];

  const applyHotReload = createConfigHotReloadAppliedHandler({
    setKeybindings: () => calls.push('set:keybindings'),
    setSessionBindings: () => calls.push('set:session-bindings'),
    refreshGlobalAndOverlayShortcuts: () => calls.push('refresh:shortcuts'),
    setSecondarySubMode: () => calls.push('set:secondary'),
    broadcastToOverlayWindows: (channel) => calls.push(`broadcast:${channel}`),
    applyAnkiRuntimeConfigPatch: (patch) => {
      calls.push('anki:patch');
      ankiPatches.push(patch);
    },
    invalidateTokenizationCache: () => calls.push('invalidate:tokens'),
    refreshSubtitlePrefetch: () => calls.push('refresh:prefetch'),
    refreshCurrentSubtitle: () => calls.push('refresh:subtitle'),
    setLogLevel: (level) => calls.push(`log:${level}`),
    setLogRotation: (rotation) => calls.push(`rotation:${rotation}`),
    setLogFileToggles: (files) => calls.push(`files:${files.mpv}`),
  });

  applyHotReload(
    {
      hotReloadFields: [
        'ankiConnect.behavior.autoUpdateNewCards',
        'ankiConnect.knownWords.highlightEnabled',
        'ankiConnect.knownWords.refreshMinutes',
        'ankiConnect.knownWords.decks',
        'ankiConnect.nPlusOne.enabled',
        'ankiConnect.nPlusOne.minSentenceWords',
        'ankiConnect.fields.word',
        'ankiConnect.fields.audio',
        'ankiConnect.fields.image',
        'ankiConnect.fields.sentence',
        'ankiConnect.fields.miscInfo',
        'ankiConnect.isLapis.sentenceCardModel',
        'ankiConnect.isKiku.fieldGrouping',
        'logging.level',
        'logging.rotation',
        'logging.files',
      ],
      restartRequiredFields: [],
    },
    config,
  );

  assert.deepEqual(ankiPatches, [
    {
      behavior: { autoUpdateNewCards: false },
      knownWords: config.ankiConnect.knownWords,
      nPlusOne: config.ankiConnect.nPlusOne,
      fields: {
        word: 'Expression',
        audio: 'SentenceAudioCustom',
        image: 'ScreenshotCustom',
        sentence: 'SentenceCustom',
        miscInfo: 'MiscInfoCustom',
      },
      isLapis: { sentenceCardModel: 'Sentence Card Custom' },
      isKiku: { fieldGrouping: 'manual' },
    },
  ]);
  assert.ok(calls.includes('invalidate:tokens'));
  assert.ok(calls.includes('refresh:prefetch'));
  assert.ok(calls.includes('refresh:subtitle'));
  assert.ok(calls.includes('log:debug'));
  assert.ok(calls.includes('rotation:14'));
  assert.ok(calls.includes('files:true'));
  assert.ok(calls.includes('broadcast:config:hot-reload'));
});

test('buildConfigHotReloadPayload includes independent primary subtitle mode', () => {
  const config = deepCloneConfig(DEFAULT_CONFIG);
  config.subtitleStyle.primaryDefaultMode = 'hover';
  config.secondarySub.defaultMode = 'hidden';

  const payload = buildConfigHotReloadPayload(config);

  assert.equal(payload.primarySubMode, 'hover');
  assert.equal(payload.secondarySubMode, 'hidden');
});

test('buildConfigHotReloadPayload reflects added, removed, and remapped session bindings', () => {
  const config = deepCloneConfig(DEFAULT_CONFIG);
  config.stats.markWatchedKey = 'Ctrl+Shift+KeyW';
  config.shortcuts.openJimaku = null;
  config.keybindings = [
    { key: 'KeyF', command: null },
    { key: 'Ctrl+Alt+KeyM', command: ['show-text', 'custom'] },
  ];

  const payload = buildConfigHotReloadPayload(config);

  assert.equal(
    payload.sessionBindings.some(
      (binding) =>
        binding.sourcePath === 'stats.markWatchedKey' &&
        binding.originalKey === 'Ctrl+Shift+KeyW' &&
        binding.actionType === 'session-action' &&
        binding.actionId === 'markWatched',
    ),
    true,
  );
  assert.equal(
    payload.sessionBindings.some(
      (binding) =>
        binding.originalKey === 'Ctrl+Alt+KeyM' &&
        binding.actionType === 'mpv-command' &&
        binding.command.join(' ') === 'show-text custom',
    ),
    true,
  );
  assert.equal(
    payload.sessionBindings.some((binding) => binding.originalKey === 'KeyF'),
    false,
  );
  assert.equal(
    payload.sessionBindings.some(
      (binding) => binding.actionType === 'session-action' && binding.actionId === 'openJimaku',
    ),
    false,
  );
});

test('createConfigHotReloadAppliedHandler skips optional effects when no hot fields', () => {
  const config = deepCloneConfig(DEFAULT_CONFIG);
  const calls: string[] = [];

  const applyHotReload = createConfigHotReloadAppliedHandler({
    setKeybindings: () => calls.push('set:keybindings'),
    setSessionBindings: () => calls.push('set:session-bindings'),
    refreshGlobalAndOverlayShortcuts: () => calls.push('refresh:shortcuts'),
    setSecondarySubMode: () => calls.push('set:secondary'),
    broadcastToOverlayWindows: (channel) => calls.push(`broadcast:${channel}`),
    applyAnkiRuntimeConfigPatch: () => calls.push('anki:patch'),
  });

  applyHotReload(
    {
      hotReloadFields: [],
      restartRequiredFields: [],
    },
    config,
  );

  assert.deepEqual(calls, ['set:keybindings', 'set:session-bindings']);
});

test('createConfigHotReloadAppliedHandler forwards compiled session-binding warnings', () => {
  const config = deepCloneConfig(DEFAULT_CONFIG);
  config.shortcuts.openSessionHelp = 'Ctrl+?';
  const warnings: string[][] = [];

  const applyHotReload = createConfigHotReloadAppliedHandler({
    setKeybindings: () => {},
    setSessionBindings: (_sessionBindings, sessionBindingWarnings) => {
      warnings.push(sessionBindingWarnings.map((warning) => warning.message));
    },
    refreshGlobalAndOverlayShortcuts: () => {},
    setSecondarySubMode: () => {},
    broadcastToOverlayWindows: () => {},
    applyAnkiRuntimeConfigPatch: () => {},
  });

  applyHotReload(
    {
      hotReloadFields: ['shortcuts'],
      restartRequiredFields: [],
    },
    config,
  );

  assert.equal(warnings.length, 1);
  assert.ok(warnings[0]?.some((message) => message.includes('Unsupported accelerator key token')));
});

test('createConfigHotReloadMessageHandler mirrors message to OSD and desktop notification', () => {
  const calls: string[] = [];
  const handleMessage = createConfigHotReloadMessageHandler({
    showMpvOsd: (message) => calls.push(`osd:${message}`),
    showDesktopNotification: (title, options) => calls.push(`notify:${title}:${options.body}`),
  });

  handleMessage('Config reload failed');
  assert.deepEqual(calls, ['osd:Config reload failed', 'notify:SubMiner:Config reload failed']);
});

test('buildRestartRequiredConfigMessage formats changed fields', () => {
  assert.equal(
    buildRestartRequiredConfigMessage(['websocket', 'subtitleStyle']),
    'Config updated; restart required for: websocket, subtitleStyle',
  );
});
