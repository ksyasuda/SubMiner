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
  assert.equal(field('secondarySub.defaultMode').category, 'behavior');
  assert.equal(field('subtitlePosition.yPercent').label, 'Subtitle Position');
});

test('settings registry groups annotation display fields by config group', () => {
  assert.equal(field('ankiConnect.knownWords.highlightEnabled').section, 'Annotation Display');
  assert.equal(field('ankiConnect.knownWords.highlightEnabled').subsection, 'Known Words');
  assert.equal(field('subtitleStyle.knownWordColor').subsection, 'Known Words');
  assert.equal(field('subtitleStyle.nPlusOneColor').subsection, 'N+1');
  assert.equal(field('subtitleStyle.enableJlpt').subsection, 'JLPT');
  assert.equal(field('subtitleStyle.jlptColors.N1').control, 'color');
});

test('settings registry exposes specialized controls for config-assisted inputs', () => {
  assert.equal(field('ankiConnect.knownWords.decks').control, 'known-words-decks');
  assert.equal(field('ankiConnect.isLapis.sentenceCardModel').control, 'anki-note-type');
  assert.equal(field('ankiConnect.fields.word').control, 'anki-field');
  assert.equal(field('keybindings').control, 'mpv-keybindings');
  assert.equal(field('shortcuts.copySubtitle').control, 'keyboard-shortcut');
  assert.equal(field('subtitleSidebar.toggleKey').control, 'key-code');
  assert.equal(field('stats.toggleKey').control, 'key-code');
  assert.equal(field('discordPresence.presenceStyle').control, 'select');
});

test('settings registry puts feature toggles first, then other toggles alphabetically', () => {
  const ankiConnect = fields.filter((candidate) => candidate.section === 'AnkiConnect');
  assert.equal(ankiConnect[0]?.configPath, 'ankiConnect.enabled');
  assert.ok(
    ankiConnect.findIndex((candidate) => candidate.configPath === 'ankiConnect.enabled') <
      ankiConnect.findIndex((candidate) => candidate.configPath === 'ankiConnect.pollingRate'),
  );

  const kikuLapis = fields.filter(
    (candidate) => candidate.section === 'Kiku Features And Lapis Features',
  );
  assert.deepEqual(
    kikuLapis.slice(0, 2).map((candidate) => candidate.configPath),
    ['ankiConnect.isLapis.enabled', 'ankiConnect.isKiku.enabled'],
  );
});

test('settings registry hides app-managed and inactive config surfaces', () => {
  const paths = new Set(fields.map((candidate) => candidate.configPath));
  for (const hiddenPath of [
    'controller.bindings',
    'controller.preferredGamepadId',
    'controller.preferredGamepadLabel',
    'controller.profiles',
    'youtubeSubgen.whisperBin',
    'jellyfin.clientVersion',
    'jellyfin.defaultLibraryId',
    'jellyfin.deviceId',
    'jellyfin.clientName',
  ]) {
    assert.equal(paths.has(hiddenPath), false, `${hiddenPath} should be hidden`);
  }
  assert.equal(field('anilist.characterDictionary.enabled').section, 'Character Dictionary');
});
