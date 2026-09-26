import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import test from 'node:test';
import { pathToFileURL } from 'node:url';
import vm from 'node:vm';
import { buildHachidoriAnkiHints, HACHIDORI_ANKI_SETTINGS_SCRIPT } from './hachidori-anki-settings';

const extensionPath = path.resolve(__dirname, '../../../../vendor/hachidori/extension');
const config = {
  tags: ['SubMiner', 'Japanese'],
  fields: { word: 'Term', sentence: 'Context', wordAudio: 'Pronunciation', image: 'Image' },
};
const modelFields = ['Term', 'Reading', 'Definition', 'Context', 'Pronunciation', 'Image'];

async function harness(
  anki: unknown = {},
  models: Record<string, string[]> = { Japanese: modelFields },
) {
  const templates: unknown = await import(
    pathToFileURL(path.join(extensionPath, 'anki-templates.js')).href
  );
  const setup: unknown = await import(
    pathToFileURL(path.join(extensionPath, 'anki-setup.js')).href
  );
  const context = vm.createContext({
    __templates: templates,
    __setup: setup,
    __initialAnki: anki,
    __models: models,
    structuredClone,
    URL,
  });
  vm.runInContext('globalThis.window = globalThis', context);
  vm.runInContext(await readFile(path.join(extensionPath, 'reader-options.js'), 'utf8'), context);
  vm.runInContext(
    `
    let options = { ...HDReaderOptions.normaliseOptions({ anki: __initialAnki }), revision: 1 };
    let writes = 0, online = true, race = false, proxy = null, failWrites = false;
    const readOptions = async () => structuredClone(options);
    globalThis.__subminerSetAnkiProxyUrl = async value => { const old = proxy; proxy = value; return old; };
    const send = async (type, request, target) => {
      if (type !== 'hd_options_write' || target !== 'hoshidicts-worker') throw Error('Unexpected request');
      if (failWrites) throw Error('offline');
      if (race) {
        race = false;
        options.anki.templates[0].tags = options.anki.tags = ['User edit'];
        options.revision++;
      }
      if (request.baseRevision !== options.revision) throw Error('conflict');
      options = { ...HDReaderOptions.normaliseOptions({ ...options, ...request.options }), revision: options.revision + 1 };
      writes++;
    };
    const __gateway = { createAnkiGateway: () => ({ discover: async ({ model }) => ({
      connected: online, model, models: Object.keys(__models), decks: ['Mining'],
      fields: __models[model] || [], errors: online ? [] : ['offline'],
    }) }) };
  `,
    context,
  );
  await vm.runInContext(
    HACHIDORI_ANKI_SETTINGS_SCRIPT.replace("await import('./anki-templates.js')", '__templates')
      .replace("await import('./anki-setup.js')", '__setup')
      .replace("await import('./anki.js')", '__gateway'),
    context,
  );
  const run = async (script: string): Promise<unknown> =>
    structuredClone(await vm.runInContext(script, context));
  const sync = () =>
    run(
      `__subminerSyncAnkiSettings(${JSON.stringify({
        server: 'http://127.0.0.1:8766',
        deck: 'Mining',
        forceOverride: true,
        hints: buildHachidoriAnkiHints(config),
      })})`,
    );
  return { run, sync };
}

test('fresh Hachidori settings inherit deck, tags, a unique model and configured fields', async () => {
  const h = await harness();
  assert.deepEqual(await h.sync(), { updated: true, matched: true, pending: false });
  assert.deepEqual(
    await h.run('[options.anki.url, options.anki.deck, options.anki.model, options.anki.tags]'),
    ['http://127.0.0.1:8766', 'Mining', 'Japanese', config.tags],
  );
  assert.deepEqual(
    await h.run(
      'Object.fromEntries(Object.entries(options.anki.fieldTemplates).map(([key, row]) => [key, row.value]))',
    ),
    {
      Term: '{expression}',
      Reading: '{reading}',
      Definition: '{definition}',
      Context: '{sentence}',
      Pronunciation: '{audio}',
      Image: '{screenshot}',
    },
  );
  await h.sync();
  assert.equal(await h.run('writes'), 1);
});

// SubMiner's new-card polling only watches ankiConnect.deck, so the first
// template's deck follows it the way Yomitan's term card deck does.
test('moves the first template to the SubMiner deck and preserves the rest of custom templates', async () => {
  const h = await harness({
    templates: [
      {
        id: 'default',
        name: 'Custom',
        deck: 'Own deck',
        model: 'Japanese',
        tags: ['own'],
        fieldTemplates: {
          Term: { value: '{reading}', overwriteMode: 'overwrite' },
          Context: { value: '', overwriteMode: 'coalesce' },
        },
      },
      { id: 'second', name: 'Second', deck: 'Other', model: 'Other', tags: [] },
    ],
  });
  const before = (await h.run('options.anki.templates')) as Array<Record<string, unknown>>;
  await h.sync();
  assert.deepEqual(await h.run('options.anki.templates'), [
    { ...before[0], deck: 'Mining' },
    ...before.slice(1),
  ]);
  assert.equal(await h.run('options.anki.deck'), 'Mining');
});

test('leaves an ambiguous model unset and retries discovery after Anki reconnects', async () => {
  const h = await harness({}, { Japanese: modelFields, Second: modelFields });
  await h.run('online = false');
  assert.deepEqual(await h.sync(), { updated: true, matched: false, pending: true });
  assert.equal(await h.run('options.anki.deck'), 'Mining');
  await h.run('online = true');
  await h.sync();
  assert.equal(await h.run('options.anki.model'), '');
  await h.run('delete __models.Second');
  await h.sync();
  assert.equal(await h.run('options.anki.model'), 'Japanese');
});

test('fills missing basic mappings only with fields belonging to the selected model', async () => {
  const h = await harness(
    { model: 'Japanese', fields: { expression: 'Reading' } },
    { Japanese: ['Term', 'Reading', 'Context'] },
  );
  await h.sync();
  assert.deepEqual(
    await h.run(
      '[options.anki.fields.expression, options.anki.fields.sentence, options.anki.fields.audio, options.anki.fieldTemplates]',
    ),
    ['Reading', 'Context', '', null],
  );
});

test('re-reads concurrent settings edits before retrying its revisioned write', async () => {
  const h = await harness();
  await h.run('race = true');
  await h.sync();
  assert.deepEqual(await h.run('options.anki.tags'), ['User edit']);
  assert.equal(await h.run('options.anki.model'), 'Japanese');
});

test('restores the previous proxy marker when every settings write fails', async () => {
  const h = await harness();
  await h.run(`proxy = 'http://127.0.0.1:9000'; failWrites = true`);
  await assert.rejects(h.sync(), /offline/);
  assert.equal(await h.run('proxy'), 'http://127.0.0.1:9000');
});

test('keeps sentence audio out of captured-audio settings and uses wordAudio first', () => {
  const hints = buildHachidoriAnkiHints({
    fields: { audio: 'SentenceAudio', wordAudio: 'WordAudio' },
  });
  assert.equal(hints.fields.audio, 'WordAudio');
  assert.equal('captureAudio' in hints.fields, false);
});
