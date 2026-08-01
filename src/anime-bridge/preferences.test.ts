import test from 'node:test';
import assert from 'node:assert/strict';
import {
  applyPreferenceValue,
  isSecretPreference,
  parsePreferences,
  type SourcePreferenceView,
} from './preferences';
import type { BridgePreference } from './types';

// Shapes taken from the real Jellyfin extension's preferencesAnime response.
const RAW: BridgePreference[] = [
  {
    key: 'host_url',
    editTextPreference: {
      title: 'Address',
      summary: 'The server address',
      value: '',
      text: '',
    },
  },
  {
    key: 'password',
    editTextPreference: { title: 'Password', summary: 'The user account password', value: '' },
  },
  {
    key: 'pref_quality',
    listPreference: {
      title: 'Preferred quality',
      summary: 'Preferred quality.',
      valueIndex: 0,
      entries: ['Source', '20 Mbps'],
      entryValues: ['source', '20000000'],
    },
  },
  {
    key: 'pref_episode_details_key',
    multiSelectListPreference: {
      title: 'Additional details for episodes',
      values: [],
      entries: ['Overview', 'Runtime'],
      entryValues: ['overview', 'runtime'],
    },
  },
  {
    key: 'pref_trust_cert',
    switchPreferenceCompat: { title: 'Trust certificate', value: false },
  },
];

function view(views: SourcePreferenceView[], key: string): SourcePreferenceView {
  const found = views.find((candidate) => candidate.key === key);
  assert.ok(found, `missing preference ${key}`);
  return found;
}

test('parsePreferences flattens each widget type', () => {
  const views = parsePreferences(RAW);
  assert.equal(views.length, 5);

  assert.equal(view(views, 'host_url').kind, 'text');
  assert.equal(view(views, 'host_url').title, 'Address');
  assert.equal(view(views, 'host_url').value, '');

  const quality = view(views, 'pref_quality');
  assert.equal(quality.kind, 'list');
  // valueIndex 0 resolves through entryValues, not entries.
  assert.equal(quality.value, 'source');
  assert.deepEqual(quality.entries, ['Source', '20 Mbps']);

  assert.deepEqual(view(views, 'pref_episode_details_key').value, []);
  assert.equal(view(views, 'pref_trust_cert').value, false);
});

test('parsePreferences skips the bridge context entry and unknown widgets', () => {
  const views = parsePreferences([
    { key: '__mangatan_bridge_context__', sourceId: '1' },
    { key: 'mystery', someFutureWidget: { title: 'X' } },
    ...RAW.slice(0, 1),
  ]);
  assert.deepEqual(
    views.map((v) => v.key),
    ['host_url'],
  );
});

test('a list preference with no selection reads as empty', () => {
  const views = parsePreferences([
    {
      key: 'library_pref',
      listPreference: {
        title: 'Select media library',
        valueIndex: -1,
        entries: [],
        entryValues: [],
      },
    },
  ]);
  assert.equal(view(views, 'library_pref').value, '');
});

test('applyPreferenceValue writes text into both value and text', () => {
  const updated = applyPreferenceValue(RAW, 'host_url', 'https://media.example');
  const body = updated.find((e) => e.key === 'host_url')!.editTextPreference as Record<
    string,
    unknown
  >;
  assert.equal(body.value, 'https://media.example');
  assert.equal(body.text, 'https://media.example');
  // Other entries are untouched.
  assert.equal(parsePreferences(updated).length, RAW.length);
});

test('applyPreferenceValue moves a list preference by entry value', () => {
  const updated = applyPreferenceValue(RAW, 'pref_quality', '20000000');
  const body = updated.find((e) => e.key === 'pref_quality')!.listPreference as Record<
    string,
    unknown
  >;
  assert.equal(body.valueIndex, 1);
  assert.equal(parsePreferences(updated).find((v) => v.key === 'pref_quality')?.value, '20000000');
});

test('applyPreferenceValue handles multi-select and switch widgets', () => {
  const multi = applyPreferenceValue(RAW, 'pref_episode_details_key', ['overview']);
  assert.deepEqual(
    parsePreferences(multi).find((v) => v.key === 'pref_episode_details_key')?.value,
    ['overview'],
  );

  const toggled = applyPreferenceValue(RAW, 'pref_trust_cert', true);
  assert.equal(parsePreferences(toggled).find((v) => v.key === 'pref_trust_cert')?.value, true);
});

test('applyPreferenceValue leaves unknown keys alone', () => {
  assert.deepEqual(applyPreferenceValue(RAW, 'not-a-key', 'x'), RAW);
});

test('secrets are recognised by key or title', () => {
  const views = parsePreferences(RAW);
  assert.equal(isSecretPreference(view(views, 'password')), true);
  assert.equal(isSecretPreference(view(views, 'host_url')), false);
});
