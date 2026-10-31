import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createCharacterDictionaryAssembly,
  type CharacterDictionaryAssemblyDeps,
} from './character-dictionary';

// Every callback throws, so a test fails if the assembly calls a dep it should only call later.
function makeDeps(): CharacterDictionaryAssemblyDeps {
  const lazy = (name: string) => (): never => {
    throw new Error(`${name} called unexpectedly`);
  };
  return {
    userDataPath: '/tmp/subminer-user-data',
    appState: {
      currentMediaPath: null,
      mpvClient: null,
      currentMediaTitle: null,
      mecabTokenizer: null,
      yomitanParserWindow: null,
      currentSubText: '',
    },
    getConfig: lazy('getConfig'),
    logger: { info: lazy('info'), warn: lazy('warn'), error: lazy('error') },
    startupOsdSequencer: {
      notifyCharacterDictionaryStatus: lazy('notifyCharacterDictionaryStatus'),
    },
    resolveMediaPathForJimaku: lazy('resolveMediaPathForJimaku'),
    getYomitanProfilePolicy: lazy('getYomitanProfilePolicy'),
    isYoutubePlaybackActive: lazy('isYoutubePlaybackActive'),
    waitForMediaTokenizationReady: lazy('waitForMediaTokenizationReady'),
    ensureYomitanExtensionLoaded: lazy('ensureYomitanExtensionLoaded'),
    getYomitanParserRuntimeDeps: lazy('getYomitanParserRuntimeDeps'),
    getOverlayNotifications: lazy('getOverlayNotifications'),
    showMpvOsd: lazy('showMpvOsd'),
    showDesktopNotification: lazy('showDesktopNotification'),
    invalidateTokenizationCache: lazy('invalidateTokenizationCache'),
    refreshCurrentSubtitle: lazy('refreshCurrentSubtitle'),
    refreshSubtitlePrefetch: lazy('refreshSubtitlePrefetch'),
    refreshCurrentSubtitleAnnotations: lazy('refreshCurrentSubtitleAnnotations'),
  };
}

// main.ts builds the assembly before most siblings exist, so construction must stay lazy.
test('createCharacterDictionaryAssembly does not call its deps while constructing', () => {
  const characterDictionary = createCharacterDictionaryAssembly(makeDeps());
  assert.equal(characterDictionary.autoSync.getCurrentMediaId(), null);
});
