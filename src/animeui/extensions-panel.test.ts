import assert from 'node:assert/strict';
import test from 'node:test';
import type { AvailableExtension, InstalledExtensionView } from '../types/anime-browser';
import { getExtensionUpdateState, summarizeExtensionUpdates } from './extensions-panel';

const installed = {
  pkg: 'pkg.example',
  name: 'Example',
  langs: ['en'],
  sourceCount: 1,
  sources: [{ id: 'pkg.example:1', name: 'Example' }],
  versionCode: 12,
  error: null,
} satisfies InstalledExtensionView;

const offered = {
  pkg: 'pkg.example',
  name: 'Example',
  lang: 'en',
  version: '1.2.0',
  versionCode: 12,
  nsfw: false,
  repoUrl: 'https://repo.example/index.json',
  iconUrl: 'https://repo.example/icon.png',
  sourceNames: ['Example'],
  installed: true,
} satisfies AvailableExtension;

test('extension update state only enables a strictly newer repository build', () => {
  assert.equal(getExtensionUpdateState(installed, { ...offered, versionCode: 13 }), 'available');
  assert.equal(getExtensionUpdateState(installed, offered), 'current');
  assert.equal(getExtensionUpdateState(installed, { ...offered, versionCode: 11 }), 'current');
});

test('extension update state does not offer unverifiable updates', () => {
  assert.equal(getExtensionUpdateState({ ...installed, versionCode: null }, offered), 'unknown');
  assert.equal(getExtensionUpdateState(installed, undefined), 'unavailable');
});

test('bulk update summary distinguishes waiting, current, and unverifiable states', () => {
  assert.deepEqual(summarizeExtensionUpdates(['current', 'available', 'available']), {
    kind: 'available',
    count: 2,
  });
  assert.deepEqual(summarizeExtensionUpdates(['current', 'current']), { kind: 'current' });
  assert.deepEqual(summarizeExtensionUpdates(['current', 'unknown']), { kind: 'none' });
  assert.deepEqual(summarizeExtensionUpdates([]), { kind: 'none' });
});
