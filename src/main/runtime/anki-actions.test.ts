import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createMarkLastCardAsAudioCardHandler,
  createMineSentenceCardHandler,
  createRefreshKnownWordCacheHandler,
  createTriggerFieldGroupingHandler,
  createUpdateLastCardFromClipboardHandler,
} from './anki-actions';

test('update last card handler forwards integration/clipboard/osd deps', async () => {
  const calls: string[] = [];
  const integration = {};
  const updateLastCard = createUpdateLastCardFromClipboardHandler({
    getAnkiIntegration: () => integration,
    readClipboardText: () => 'clipboard-value',
    showMpvOsd: (text) => calls.push(`osd:${text}`),
    updateLastCardFromClipboardCore: async (options) => {
      assert.equal(options.ankiIntegration, integration);
      assert.equal(options.readClipboardText(), 'clipboard-value');
      options.showMpvOsd('ok');
      calls.push('core');
    },
  });

  await updateLastCard();
  assert.deepEqual(calls, ['osd:ok', 'core']);
});

test('refresh known word cache handler throws when Anki integration missing', async () => {
  const refresh = createRefreshKnownWordCacheHandler({
    getAnkiIntegration: () => null,
    missingIntegrationMessage: 'AnkiConnect integration not enabled',
  });

  await assert.rejects(() => refresh(), /AnkiConnect integration not enabled/);
});

test('trigger and mark handlers delegate to core services', async () => {
  const calls: string[] = [];
  const integration = {};
  const triggerFieldGrouping = createTriggerFieldGroupingHandler({
    getAnkiIntegration: () => integration,
    showMpvOsd: (text) => calls.push(`osd:${text}`),
    triggerFieldGroupingCore: async (options) => {
      assert.equal(options.ankiIntegration, integration);
      options.showMpvOsd('group');
      calls.push('group-core');
    },
  });
  const markAudio = createMarkLastCardAsAudioCardHandler({
    getAnkiIntegration: () => integration,
    showMpvOsd: (text) => calls.push(`osd:${text}`),
    markLastCardAsAudioCardCore: async (options) => {
      assert.equal(options.ankiIntegration, integration);
      options.showMpvOsd('mark');
      calls.push('mark-core');
    },
  });

  await triggerFieldGrouping();
  await markAudio();
  assert.deepEqual(calls, ['osd:group', 'group-core', 'osd:mark', 'mark-core']);
});

test('mine sentence handler records mined cards only when core returns true', async () => {
  const calls: string[] = [];
  const integration = {};
  const mpvClient = {};
  let created = false;
  const mineSentenceCard = createMineSentenceCardHandler({
    getAnkiIntegration: () => integration,
    getMpvClient: () => mpvClient,
    showMpvOsd: (text) => calls.push(`osd:${text}`),
    mineSentenceCardCore: async (options) => {
      assert.equal(options.ankiIntegration, integration);
      assert.equal(options.mpvClient, mpvClient);
      options.showMpvOsd('mine');
      return created;
    },
    recordCardsMined: (count) => calls.push(`cards:${count}`),
  });

  created = false;
  await mineSentenceCard();
  created = true;
  await mineSentenceCard();
  assert.deepEqual(calls, ['osd:mine', 'osd:mine', 'cards:1']);
});
