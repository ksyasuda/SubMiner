import test from 'node:test';
import assert from 'node:assert/strict';
import { DEFAULT_CONFIG } from '../definitions';
import { buildConfigSettingsRegistry } from './registry';

const fields = buildConfigSettingsRegistry(DEFAULT_CONFIG);

function field(path: string) {
  const match = fields.find((candidate) => candidate.configPath === path);
  assert.ok(match, `missing settings field: ${path}`);
  return match;
}

test('settings registry splits viewing into appearance and behavior categories', () => {
  assert.equal(field('subtitleStyle.fontSize').category, 'appearance');
  assert.equal(field('subtitleStyle.primaryDefaultMode').category, 'behavior');
  assert.equal(field('subtitleStyle.primaryDefaultMode').section, 'Subtitle Behavior');
  assert.equal(field('subtitleStyle.primaryVisibleOnYomitanPopup').category, 'behavior');
  assert.equal(field('subtitleStyle.primaryVisibleOnYomitanPopup').section, 'Subtitle Behavior');
  assert.equal(field('secondarySub.defaultMode').category, 'behavior');
  assert.equal(field('subtitlePosition.yPercent').label, 'Subtitle Position');
  assert.equal(field('subtitleStyle.frequencyDictionary.mode').label, 'Frequency Mode');
  assert.equal(field('auto_start_overlay').category, 'behavior');
  assert.equal(field('auto_start_overlay').section, 'Playback Behavior');
  assert.equal(field('youtube.primarySubLanguages').category, 'behavior');
  assert.equal(field('youtube.primarySubLanguages').section, 'YouTube Playback Settings');
  assert.equal(field('mpv.launchMode').category, 'behavior');
  assert.equal(field('mpv.launchMode').section, 'mpv Playback');
  assert.equal(field('mpv.profile').category, 'behavior');
  assert.equal(field('mpv.profile').section, 'mpv Playback');
  assert.ok(
    fields.findIndex((candidate) => candidate.configPath === 'subtitleStyle.primaryDefaultMode') <
      fields.findIndex(
        (candidate) => candidate.configPath === 'subtitleStyle.primaryVisibleOnYomitanPopup',
      ),
  );
  assert.ok(
    fields.findIndex(
      (candidate) => candidate.configPath === 'subtitleStyle.primaryVisibleOnYomitanPopup',
    ) < fields.findIndex((candidate) => candidate.configPath === 'secondarySub.defaultMode'),
  );
});

test('settings registry groups playback startup controls under playback behavior', () => {
  for (const path of [
    'subtitleStyle.autoPauseVideoOnHover',
    'subtitleStyle.autoPauseVideoOnYomitanPopup',
    'subtitleSidebar.pauseVideoOnHover',
    'mpv.autoStartSubMiner',
    'auto_start_overlay',
    'mpv.pauseUntilOverlayReady',
  ]) {
    assert.equal(field(path).category, 'behavior', path);
    assert.equal(field(path).section, 'Playback Behavior', path);
  }
});

test('settings registry moves AniSkip button key into input shortcuts and hot reload', () => {
  assert.equal(field('mpv.aniskipButtonKey').category, 'input');
  assert.equal(field('mpv.aniskipButtonKey').section, 'Overlay Shortcuts');
  assert.equal(field('mpv.aniskipButtonKey').subsection, 'Playback');
  assert.equal(field('mpv.aniskipButtonKey').control, 'mpv-key');
  assert.equal(field('mpv.aniskipButtonKey').restartBehavior, 'hot-reload');
});

test('settings registry exposes character dictionary panel shortcuts dynamically', () => {
  assert.equal(
    field('shortcuts.openCharacterDictionaryManager').label,
    'Open Character Dictionary Manager',
  );
  assert.equal(field('shortcuts.openCharacterDictionaryManager').subsection, 'Open Panels');
});

test('settings registry hides removed modal-only fields', () => {
  for (const path of [
    'shortcuts.multiCopyTimeoutMs',
    'anilist.characterDictionary.profileScope',
    'jellyfin.directPlayContainers',
  ]) {
    assert.equal(
      fields.some((candidate) => candidate.configPath === path),
      false,
      path,
    );
  }
});

test('settings registry orders websocket server immediately after annotation websocket', () => {
  const integrationSections = [
    ...new Set(
      fields
        .filter((candidate) => candidate.category === 'integrations')
        .map((candidate) => candidate.section),
    ),
  ];
  const annotationIndex = integrationSections.indexOf('Annotation WebSocket');
  assert.equal(integrationSections[annotationIndex + 1], 'WebSocket server');
});

test('settings registry explains websocket auto mode and keeps it disabled by default', () => {
  assert.equal(field('websocket.enabled').defaultValue, false);
  assert.equal(
    field('websocket.enabled').description,
    'Built-in subtitle WebSocket server mode. Auto starts the built-in server only when mpv_websocket is not detected; otherwise it defers to the plugin.',
  );
});

test('settings registry places immersion tracking after other tracking and app sections', () => {
  const trackingSections = [
    ...new Set(
      fields
        .filter((candidate) => candidate.category === 'tracking-app')
        .map((candidate) => candidate.section),
    ),
  ];
  assert.equal(trackingSections.at(-1), 'Immersion tracking');
});

test('settings registry groups annotation display fields by config group', () => {
  assert.equal(field('ankiConnect.knownWords.highlightEnabled').section, 'Annotation Display');
  assert.equal(field('ankiConnect.knownWords.highlightEnabled').subsection, 'Known Words');
  assert.equal(field('subtitleStyle.knownWordColor').subsection, 'Known Words');
  assert.equal(field('subtitleStyle.nPlusOneColor').subsection, 'N+1');
  assert.equal(field('subtitleStyle.enableJlpt').subsection, 'JLPT');
  assert.equal(field('subtitleStyle.jlptColors.N1').control, 'color');
});

test('settings registry routes known words sync, n+1, and frequency config to behavior', () => {
  assert.equal(field('ankiConnect.knownWords.addMinedWordsImmediately').category, 'behavior');
  assert.equal(field('ankiConnect.knownWords.addMinedWordsImmediately').section, 'Known Words');
  assert.equal(field('ankiConnect.knownWords.decks').category, 'behavior');
  assert.equal(field('ankiConnect.knownWords.decks').section, 'Known Words');
  assert.equal(field('ankiConnect.knownWords.matchMode').category, 'behavior');
  assert.equal(field('ankiConnect.knownWords.matchMode').section, 'Known Words');
  assert.equal(field('ankiConnect.knownWords.refreshMinutes').category, 'behavior');
  assert.equal(field('ankiConnect.knownWords.refreshMinutes').section, 'Known Words');
  assert.equal(field('ankiConnect.nPlusOne.minSentenceWords').category, 'behavior');
  assert.equal(field('ankiConnect.nPlusOne.minSentenceWords').section, 'N+1');
  assert.equal(field('subtitleStyle.frequencyDictionary.sourcePath').category, 'behavior');
  assert.equal(
    field('subtitleStyle.frequencyDictionary.sourcePath').section,
    'Frequency Highlighting',
  );
  assert.equal(field('subtitleStyle.frequencyDictionary.mode').category, 'behavior');
  assert.equal(field('subtitleStyle.frequencyDictionary.matchMode').category, 'behavior');
  assert.equal(field('subtitleStyle.frequencyDictionary.topX').category, 'behavior');
});

test('settings registry exposes mpv aniskip button as an mpv key learn control', () => {
  assert.equal(field('mpv.aniskipButtonKey').control, 'mpv-key');
});

test('settings registry exposes specialized controls for config-assisted inputs', () => {
  assert.equal(field('ankiConnect.deck').control, 'anki-deck');
  assert.equal(field('ankiConnect.knownWords.decks').control, 'known-words-decks');
  assert.equal(field('ankiConnect.isLapis.sentenceCardModel').control, 'anki-note-type');
  assert.equal(field('ankiConnect.fields.word').control, 'anki-field');
  assert.equal(field('keybindings').control, 'mpv-keybindings');
  assert.equal(field('subtitleStyle.css').control, 'css-declarations');
  assert.equal(field('subtitleStyle.secondary.css').control, 'css-declarations');
  assert.equal(field('shortcuts.copySubtitle').control, 'keyboard-shortcut');
  assert.equal(field('mpv.aniskipButtonKey').control, 'mpv-key');
  assert.equal(field('subtitleSidebar.css').control, 'css-declarations');
  assert.equal(field('stats.toggleKey').control, 'key-code');
  assert.equal(field('discordPresence.presenceStyle').control, 'select');
});

test('settings registry exposes css declaration editor for primary and secondary subtitle appearance', () => {
  const primaryVisible = fields
    .filter(
      (candidate) =>
        candidate.section === 'Primary Subtitle Appearance' && !candidate.settingsHidden,
    )
    .map((candidate) => candidate.configPath);
  const secondaryVisible = fields
    .filter(
      (candidate) =>
        candidate.section === 'Secondary Subtitle Appearance' && !candidate.settingsHidden,
    )
    .map((candidate) => candidate.configPath);

  assert.deepEqual(primaryVisible, ['subtitleStyle.css']);
  assert.deepEqual(secondaryVisible, ['subtitleStyle.secondary.css']);
  assert.equal(field('subtitleStyle.fontSize').settingsHidden, true);
  assert.equal(field('subtitleStyle.secondary.fontSize').settingsHidden, true);
  assert.equal(field('subtitleStyle.fontColor').settingsHidden, true);
  assert.equal(field('subtitleStyle.backgroundColor').settingsHidden, true);
  assert.equal(field('subtitleStyle.hoverTokenColor').settingsHidden, true);
  assert.equal(field('subtitleStyle.hoverTokenBackgroundColor').settingsHidden, true);
  assert.equal(field('subtitleStyle.paintOrder').settingsHidden, true);
  assert.equal(field('subtitleStyle.WebkitTextStroke').settingsHidden, true);
  assert.equal(field('subtitleStyle.knownWordColor').settingsHidden, false);
  assert.equal(field('subtitleStyle.nPlusOneColor').settingsHidden, false);
  assert.equal(field('subtitleStyle.nameMatchImagesEnabled').settingsHidden, false);
  assert.equal(field('subtitleStyle.nameMatchColor').settingsHidden, false);
  assert.equal(field('subtitleStyle.jlptColors.N1').settingsHidden, false);
  assert.equal(field('subtitleStyle.frequencyDictionary.singleColor').settingsHidden, false);
  assert.equal(field('subtitleStyle.frequencyDictionary.bandedColors').settingsHidden, false);
});

test('settings registry exposes css declaration editor for subtitle sidebar appearance', () => {
  const sidebarVisible = fields
    .filter(
      (candidate) =>
        candidate.section === 'Subtitle Sidebar Appearance' && !candidate.settingsHidden,
    )
    .map((candidate) => candidate.configPath);

  assert.deepEqual(sidebarVisible, ['subtitleSidebar.css']);
  assert.equal(field('subtitleSidebar.fontFamily').settingsHidden, true);
  assert.equal(field('subtitleSidebar.fontSize').settingsHidden, true);
  assert.equal(field('subtitleSidebar.textColor').settingsHidden, true);
  assert.equal(field('subtitleSidebar.backgroundColor').settingsHidden, true);
  assert.equal(field('subtitleSidebar.timestampColor').settingsHidden, true);
  assert.equal(field('subtitleSidebar.activeLineColor').settingsHidden, true);
  assert.equal(field('subtitleSidebar.activeLineBackgroundColor').settingsHidden, true);
  assert.equal(field('subtitleSidebar.hoverLineBackgroundColor').settingsHidden, true);
  assert.equal(field('subtitleSidebar.enabled').settingsHidden, false);
  assert.equal(field('subtitleSidebar.layout').settingsHidden, false);
});

test('settings registry routes playback-related integrations into integrations', () => {
  assert.equal(field('jimaku.apiBaseUrl').category, 'integrations');
  assert.equal(field('jimaku.apiBaseUrl').section, 'Jimaku');
  assert.equal(field('subsync.replace').category, 'integrations');
  assert.equal(field('subsync.replace').section, 'Subtitle Sync');
});

test('settings registry puts feature toggles first, then other toggles alphabetically', () => {
  const ankiConnect = fields.filter((candidate) => candidate.section === 'AnkiConnect');
  assert.equal(ankiConnect[0]?.configPath, 'ankiConnect.enabled');
  assert.equal(ankiConnect[1]?.configPath, 'ankiConnect.deck');
  assert.ok(
    ankiConnect.findIndex((candidate) => candidate.configPath === 'ankiConnect.enabled') <
      ankiConnect.findIndex((candidate) => candidate.configPath === 'ankiConnect.pollingRate'),
  );
  assert.ok(
    fields.findIndex((candidate) => candidate.section === 'AnkiConnect') <
      fields.findIndex((candidate) => candidate.section === 'AnkiConnect Proxy'),
  );
  const miningSections = [
    ...new Set(
      fields
        .filter((candidate) => candidate.category === 'mining-anki')
        .map((candidate) => candidate.section),
    ),
  ];
  assert.equal(miningSections[0], 'AnkiConnect');

  const kikuLapis = fields.filter((candidate) => candidate.section === 'Kiku/Lapis Features');
  assert.deepEqual(
    kikuLapis.slice(0, 2).map((candidate) => candidate.configPath),
    ['ankiConnect.isLapis.enabled', 'ankiConnect.isKiku.enabled'],
  );
});

test('settings registry hides app-managed and inactive config surfaces', () => {
  const paths = new Set(fields.map((candidate) => candidate.configPath));
  for (const hiddenPath of [
    'ai.enabled',
    'ai.apiKey',
    'ai.apiKeyCommand',
    'ai.model',
    'ai.baseUrl',
    'ai.systemPrompt',
    'ai.requestTimeoutMs',
    'ankiConnect.ai.enabled',
    'ankiConnect.ai.model',
    'ankiConnect.ai.systemPrompt',
    'ankiConnect.fields.translation',
    'controller.bindings',
    'controller.preferredGamepadId',
    'controller.preferredGamepadLabel',
    'controller.profiles',
    'youtubeSubgen.whisperBin',
    'jellyfin.defaultLibraryId',
    'subtitleSidebar.toggleKey',
    'jellyfin.recentServers',
  ]) {
    assert.equal(paths.has(hiddenPath), false, `${hiddenPath} should be hidden`);
  }
  assert.equal(paths.has('anilist.characterDictionary.enabled'), false);
});

test('settings registry marks safe live config paths as hot-reloadable', () => {
  for (const path of [
    'mpv.aniskipButtonKey',
    'stats.toggleKey',
    'stats.markWatchedKey',
    'logging.level',
    'logging.rotation',
    'logging.files.app',
    'logging.files.launcher',
    'logging.files.mpv',
    'youtube.primarySubLanguages',
    'jimaku.apiBaseUrl',
    'jimaku.languagePreference',
    'jimaku.maxEntryResults',
    'subsync.replace',
    'ankiConnect.behavior.autoUpdateNewCards',
    'ankiConnect.deck',
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
  ]) {
    assert.equal(field(path).restartBehavior, 'hot-reload', path);
  }
});

test('settings registry does not expose removed subsync mode option', () => {
  const paths = new Set(fields.map((candidate) => candidate.configPath));
  assert.equal(paths.has('subsync.defaultMode'), false);
});

test('settings registry keeps unsafe config siblings restart-required', () => {
  for (const path of [
    'stats.serverPort',
    'ankiConnect.url',
    'ankiConnect.proxy.enabled',
    'mpv.socketPath',
    'mpv.profile',
    'websocket.port',
  ]) {
    assert.equal(field(path).restartBehavior, 'restart', path);
  }
});
