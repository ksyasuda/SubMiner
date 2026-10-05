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

  const kikuLapis = fields.filter(
    (candidate) => candidate.section === 'Kiku/Lapis/Senren Features',
  );
  assert.deepEqual(
    kikuLapis.slice(0, 3).map((candidate) => candidate.configPath),
    ['ankiConnect.isLapis.enabled', 'ankiConnect.isKiku.enabled', 'ankiConnect.isSenren.enabled'],
  );
});
