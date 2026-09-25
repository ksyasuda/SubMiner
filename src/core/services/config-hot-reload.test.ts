import test from 'node:test';
import assert from 'node:assert/strict';
import { DEFAULT_CONFIG, deepCloneConfig } from '../../config';
import { buildConfigSettingsRegistry, getConfigValueAtPath } from '../../config/settings/registry';
import {
  classifyConfigHotReloadDiff,
  createConfigHotReloadRuntime,
  type ConfigHotReloadRuntimeDeps,
} from './config-hot-reload';

test('every LIVE settings field is classified without a restart warning', () => {
  for (const field of buildConfigSettingsRegistry(DEFAULT_CONFIG)) {
    if (field.restartBehavior !== 'hot-reload') continue;
    const next = deepCloneConfig(DEFAULT_CONFIG);
    const segments = field.configPath.split('.');
    const leaf = segments.pop();
    assert.ok(leaf);
    const parent = segments.length ? getConfigValueAtPath(next, segments.join('.')) : next;
    assert.ok(parent && typeof parent === 'object', field.configPath);
    // The classifier compares structure; validation of field values is tested separately.
    Object.defineProperty(parent, leaf, {
      value: getConfigValueAtPath(next, field.configPath) === null ? 'changed' : null,
      enumerable: true,
    });
    const diff = classifyConfigHotReloadDiff(DEFAULT_CONFIG, next);
    assert.deepEqual(diff.restartRequiredFields, [], field.configPath);
    assert.ok(diff.hotReloadFields.length > 0, field.configPath);
  }
});

test('live notifications and subtitle generation changes preserve unrelated restart warnings', () => {
  const next = deepCloneConfig(DEFAULT_CONFIG);
  next.notifications.overlayPosition = 'top';
  next.subtitleGeneration.threads += 1;
  next.websocket.port += 1;

  const diff = classifyConfigHotReloadDiff(DEFAULT_CONFIG, next);
  assert.deepEqual(
    new Set(diff.hotReloadFields),
    new Set(['notifications.overlayPosition', 'subtitleGeneration.threads']),
  );
  assert.deepEqual(diff.restartRequiredFields, ['websocket.port']);
});

test('classifyConfigHotReloadDiff separates hot and restart-required fields', () => {
  const prev = deepCloneConfig(DEFAULT_CONFIG);
  const next = deepCloneConfig(DEFAULT_CONFIG);
  next.subtitleStyle.fontSize = prev.subtitleStyle.fontSize + 2;
  next.websocket.port = prev.websocket.port + 1;

  const diff = classifyConfigHotReloadDiff(prev, next);
  assert.deepEqual(diff.hotReloadFields, ['subtitleStyle']);
  assert.deepEqual(diff.restartRequiredFields, ['websocket.port']);
});

test('classifyConfigHotReloadDiff treats safe nested config paths as hot-reloadable', () => {
  const prev = deepCloneConfig(DEFAULT_CONFIG);
  const next = deepCloneConfig(DEFAULT_CONFIG);
  next.mpv.aniskipButtonKey = 'F8';
  next.stats.toggleKey = 'F8';
  next.stats.markWatchedKey = 'F9';
  next.logging.level = 'debug';
  next.logging.rotation = 14;
  next.logging.files.mpv = true;
  next.youtube.primarySubLanguages = ['ja', 'en'];
  next.jimaku.maxEntryResults = prev.jimaku.maxEntryResults + 1;
  next.subsync.replace = !prev.subsync.replace;
  next.ankiConnect.deck = 'Mining';
  next.ankiConnect.media.normalizeAudio = !prev.ankiConnect.media.normalizeAudio;
  next.ankiConnect.media.mirrorMpvVolume = !prev.ankiConnect.media.mirrorMpvVolume;
  next.ankiConnect.media.reviewTiming = !prev.ankiConnect.media.reviewTiming;
  next.ankiConnect.behavior.autoUpdateNewCards = !prev.ankiConnect.behavior.autoUpdateNewCards;
  next.ankiConnect.knownWords.highlightEnabled = !prev.ankiConnect.knownWords.highlightEnabled;
  next.ankiConnect.knownWords.refreshMinutes = prev.ankiConnect.knownWords.refreshMinutes + 5;
  next.ankiConnect.knownWords.addMinedWordsImmediately =
    !prev.ankiConnect.knownWords.addMinedWordsImmediately;
  next.ankiConnect.knownWords.matchMode =
    prev.ankiConnect.knownWords.matchMode === 'headword' ? 'surface' : 'headword';
  next.ankiConnect.knownWords.decks = { Anime: ['Mining'] };
  next.ankiConnect.nPlusOne.enabled = !prev.ankiConnect.nPlusOne.enabled;
  next.ankiConnect.nPlusOne.minSentenceWords = prev.ankiConnect.nPlusOne.minSentenceWords + 1;
  next.ankiConnect.fields.word = 'Vocabulary';
  next.ankiConnect.fields.audio = 'SentenceAudioCustom';
  next.ankiConnect.fields.image = 'ScreenshotCustom';
  next.ankiConnect.fields.sentence = 'SentenceCustom';
  next.ankiConnect.fields.miscInfo = 'MiscInfoCustom';
  next.ankiConnect.isLapis.sentenceCardModel = 'Sentence Card Custom';
  next.ankiConnect.isKiku.fieldGrouping =
    prev.ankiConnect.isKiku.fieldGrouping === 'auto' ? 'manual' : 'auto';

  const diff = classifyConfigHotReloadDiff(prev, next);

  assert.deepEqual(
    new Set(diff.hotReloadFields),
    new Set([
      'stats.toggleKey',
      'mpv.aniskipButtonKey',
      'stats.markWatchedKey',
      'logging.level',
      'logging.rotation',
      'logging.files',
      'youtube.primarySubLanguages',
      'jimaku.maxEntryResults',
      'subsync.replace',
      'ankiConnect.deck',
      'ankiConnect.media.normalizeAudio',
      'ankiConnect.media.mirrorMpvVolume',
      'ankiConnect.media.reviewTiming',
      'ankiConnect.behavior.autoUpdateNewCards',
      'ankiConnect.knownWords.highlightEnabled',
      'ankiConnect.knownWords.refreshMinutes',
      'ankiConnect.knownWords.addMinedWordsImmediately',
      'ankiConnect.knownWords.matchMode',
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
    ]),
  );
  assert.deepEqual(diff.restartRequiredFields, []);
});

test('classifyConfigHotReloadDiff keeps unsafe nested siblings restart-required', () => {
  const prev = deepCloneConfig(DEFAULT_CONFIG);
  const next = deepCloneConfig(DEFAULT_CONFIG);
  next.stats.serverPort = prev.stats.serverPort + 1;
  next.ankiConnect.url = 'http://127.0.0.1:9999';
  next.ankiConnect.ai.model = 'openrouter/new-model';

  const diff = classifyConfigHotReloadDiff(prev, next);

  assert.deepEqual(diff.hotReloadFields, []);
  assert.deepEqual(diff.restartRequiredFields, [
    'ankiConnect.url',
    'ankiConnect.ai.model',
    'stats.serverPort',
  ]);
});

test('config hot reload runtime debounces rapid watch events', () => {
  let watchedChangeCallback: (() => void) | null = null;
  const pendingTimers = new Map<number, () => void>();
  let nextTimerId = 1;
  let reloadCalls = 0;

  const deps: ConfigHotReloadRuntimeDeps = {
    getCurrentConfig: () => deepCloneConfig(DEFAULT_CONFIG),
    reloadConfigStrict: () => {
      reloadCalls += 1;
      return {
        ok: true,
        config: deepCloneConfig(DEFAULT_CONFIG),
        warnings: [],
        path: '/tmp/config.jsonc',
      };
    },
    watchConfigPath: (_path, onChange) => {
      watchedChangeCallback = onChange;
      return { close: () => {} };
    },
    setTimeout: (callback) => {
      const id = nextTimerId;
      nextTimerId += 1;
      pendingTimers.set(id, callback);
      return id as unknown as NodeJS.Timeout;
    },
    clearTimeout: (timeout) => {
      pendingTimers.delete(timeout as unknown as number);
    },
    debounceMs: 25,
    onHotReloadApplied: () => {},
    onRestartRequired: () => {},
    onInvalidConfig: () => {},
    onValidationWarnings: () => {},
  };

  const runtime = createConfigHotReloadRuntime(deps);
  runtime.start();
  assert.equal(reloadCalls, 1);
  if (!watchedChangeCallback) {
    throw new Error('Expected watch callback to be registered.');
  }
  const trigger = watchedChangeCallback as () => void;

  trigger();
  trigger();
  trigger();
  assert.equal(pendingTimers.size, 1);

  for (const callback of pendingTimers.values()) {
    callback();
  }
  assert.equal(reloadCalls, 2);
});

test('config hot reload runtime reports invalid config and skips apply', () => {
  const invalidMessages: string[] = [];
  let watchedChangeCallback: (() => void) | null = null;

  const runtime = createConfigHotReloadRuntime({
    getCurrentConfig: () => deepCloneConfig(DEFAULT_CONFIG),
    reloadConfigStrict: () => ({
      ok: false,
      error: 'Invalid JSON',
      path: '/tmp/config.jsonc',
    }),
    watchConfigPath: (_path, onChange) => {
      watchedChangeCallback = onChange;
      return { close: () => {} };
    },
    setTimeout: (callback) => {
      callback();
      return 1 as unknown as NodeJS.Timeout;
    },
    clearTimeout: () => {},
    debounceMs: 0,
    onHotReloadApplied: () => {
      throw new Error('Hot reload should not apply for invalid config.');
    },
    onRestartRequired: () => {
      throw new Error('Restart warning should not trigger for invalid config.');
    },
    onInvalidConfig: (message) => {
      invalidMessages.push(message);
    },
    onValidationWarnings: () => {
      throw new Error('Validation warnings should not trigger for invalid config.');
    },
  });

  runtime.start();
  assert.equal(watchedChangeCallback, null);
  assert.equal(invalidMessages.length, 1);
});

test('config hot reload runtime reports validation warnings from reload', () => {
  let watchedChangeCallback: (() => void) | null = null;
  const warningCalls: Array<{ path: string; count: number }> = [];

  const runtime = createConfigHotReloadRuntime({
    getCurrentConfig: () => deepCloneConfig(DEFAULT_CONFIG),
    reloadConfigStrict: () => ({
      ok: true,
      config: deepCloneConfig(DEFAULT_CONFIG),
      warnings: [
        {
          path: 'ankiConnect.ai',
          message: 'Expected boolean.',
          value: { enabled: true },
          fallback: false,
        },
      ],
      path: '/tmp/config.jsonc',
    }),
    watchConfigPath: (_path, onChange) => {
      watchedChangeCallback = onChange;
      return { close: () => {} };
    },
    setTimeout: (callback) => {
      callback();
      return 1 as unknown as NodeJS.Timeout;
    },
    clearTimeout: () => {},
    debounceMs: 0,
    onHotReloadApplied: () => {},
    onRestartRequired: () => {},
    onInvalidConfig: () => {},
    onValidationWarnings: (path, warnings) => {
      warningCalls.push({ path, count: warnings.length });
    },
  });

  runtime.start();
  assert.equal(warningCalls.length, 0);
  if (!watchedChangeCallback) {
    throw new Error('Expected watch callback to be registered.');
  }
  const trigger = watchedChangeCallback as () => void;
  trigger();
  assert.deepEqual(warningCalls, [{ path: '/tmp/config.jsonc', count: 1 }]);
});
