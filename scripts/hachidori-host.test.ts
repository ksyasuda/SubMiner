import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { runInNewContext } from 'node:vm';
import test from 'node:test';
import { CHARACTER_DICTIONARY_TITLE_PREFIX } from '../src/core/services/tokenizer/character-dictionary-title';
import { createAnkiGateway } from '../vendor/hachidori/extension/anki.js';

const bridge = readFileSync(
  new URL('../vendor/hachidori/extension/subminer-host.js', import.meta.url),
  'utf8',
);

function run(code: string) {
  runInNewContext(`${bridge}\n${code}`, {
    window: new EventTarget(),
    EventTarget,
    CustomEvent,
    assert,
    characterPrefix: CHARACTER_DICTIONARY_TITLE_PREFIX,
    KeyboardEvent: class {
      constructor(_type: string, options: KeyboardEventInit = {}) {
        Object.assign(this, options);
      }
    },
  });
}

test('Hachidori marks popup state independently from successful lookups', () => {
  run(`
    const events = [];
    for (const name of ['yomitan-popup-shown', 'yomitan-popup-hidden', 'subminer-yomitan-lookup']) {
      window.addEventListener(name, () => events.push(name));
    }
    const attributes = new Map();
    const host = { setAttribute: (name, value) => attributes.set(name, value) };
    SubMinerHachidori.markHost(host, true);
    assert.equal(attributes.get('data-subminer-yomitan-popup-visible'), 'true');
    assert.equal(events.join(','), '');
    SubMinerHachidori.lookup();
    SubMinerHachidori.markHost(host, false);
    assert.equal(attributes.get('data-subminer-yomitan-popup-visible'), 'false');
    assert.equal(events.join(','), 'subminer-yomitan-lookup');
  `);
});

test('Hachidori promotes character glossaries and results without overriding linguistic rank', () => {
  run(`
    const result = (id, matched, dictionary, options = {}) => ({
      matched, deinflected: matched, trace: [], preprocessorSteps: 0,
      term: { expression: matched, reading: 'reading', glossaries: [{ dictionary }] }, ...options, id,
    });
    const character = characterPrefix + ' - Current show';
    const longer = result('longer', '花子さん', 'General');
    const normal = result('normal', '花子', 'General');
    const person = result('person', '花子', character);
    const shorter = result('shorter', '花', character);
    const output = SubMinerHachidori.prioritizeCharacterResults([longer, normal, person, shorter]);
    assert.equal(output.map(entry => entry.id).join(','), 'longer,person,normal,shorter');
    const merged = result('merged', '花子', 'General');
    merged.term.glossaries.push({ dictionary: character }, { dictionary: 'Second general' });
    const [promoted] = SubMinerHachidori.prioritizeCharacterResults([merged]);
    assert.equal(promoted.term.glossaries.map(g => g.dictionary).join(','), character + ',General,Second general');
    assert.equal(merged.term.glossaries[0].dictionary, 'General');
    for (const change of [{ preprocessorSteps: 1 }, { trace: [{}] }, { deinflected: '別の語' }]) {
      const transformed = { ...person, ...change };
      assert.equal(SubMinerHachidori.prioritizeCharacterResults([normal, transformed])[0].id, 'normal');
    }
    const preferredReading = { ...normal, term: { ...normal.term, reading: 'preferred' } };
    assert.equal(SubMinerHachidori.prioritizeCharacterResults([person, preferredReading], {primaryReading: 'preferred'})[0].id, 'normal');
    const aliased = result('aliased', '花子', 'Imported character data');
    assert.equal(SubMinerHachidori.prioritizeCharacterResults([normal, aliased], {}, [{title: 'Imported character data', displayName: character}])[0].id, 'aliased');
  `);
});

test('Hachidori sends private mining metadata only to the configured SubMiner proxy', async () => {
  const requests: Array<{ url: string; body: string }> = [];
  const proxy = 'http://127.0.0.1:8766';
  const gateway = createAnkiGateway({
    readSubminerProxyUrl: async () => proxy,
    fetch: async (url: string, options: RequestInit) => {
      assert.equal(typeof options.body, 'string');
      requests.push({ url, body: String(options.body) });
      return Response.json({ result: 123, error: null });
    },
  });
  const params = {
    note: { fields: { Expression: '花子' } },
    subminerDuplicateNoteIds: [456],
    subminerEnrich: true,
  };
  await gateway.invoke('addNote', params, '', 1000, proxy);
  await gateway.invoke('addNote', params, '', 1000, 'http://127.0.0.1:8765');
  assert.deepEqual(
    requests.map((request) => JSON.parse(request.body).params),
    [params, { note: params.note }],
  );
  assert.equal(params.subminerEnrich, true);
  const direct = createAnkiGateway({
    readSubminerProxyUrl: async () => null,
    fetch: async (_url: string, options: RequestInit) => {
      assert.deepEqual(JSON.parse(String(options.body)).params, { note: params.note });
      return Response.json({ result: null, error: null });
    },
  });
  await direct.invoke('updateNoteFields', params, '', 1000, proxy);
});

test('Hachidori routes host commands, validates keyboard input and disconnects', () => {
  run(`
    const calls = [];
    const disconnect = SubMinerHachidori.connect({
      hide: () => calls.push('hide'), clear: () => calls.push('clear'),
      action: name => calls.push(name), cycleAudio: direction => calls.push(direction),
      scroll: (x, y) => calls.push(x + ':' + y),
      keydown: event => calls.push(event.key + ':' + event.ctrlKey + ':' + event.shiftKey),
    });
    const send = detail => window.dispatchEvent(new CustomEvent('subminer-yomitan-popup-command', {detail}));
    send(null);
    send({ type: 'forwardKeyDown', key: 12, modifiers: [] });
    send({ type: 'mineSelected' });
    send({ type: 'playCurrentAudio' });
    send({ type: 'scanSelectedText' });
    send({ type: 'cycleAudioSource', direction: -1 });
    send({ type: 'scrollBy', deltaX: Infinity, deltaY: 40 });
    send({ type: 'forwardKeyDown', key: 'j', code: 'KeyJ', modifiers: ['ctrl', 'shift'] });
    send({ type: 'setVisible', visible: false });
    send({ type: 'clearActiveTextSource' });
    disconnect();
    send({ type: 'mineSelected' });
    assert.equal(calls.join(','), 'addNote,playAudio,scanSelectedText,-1,0:40,j:true:true,hide,clear');
  `);
});
