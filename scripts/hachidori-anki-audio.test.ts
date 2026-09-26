import assert from 'node:assert/strict';
import test from 'node:test';
import '../vendor/hachidori/extension/reader-options.js';
import { createAnkiGateway } from '../vendor/hachidori/extension/anki.js';
import { createAnkiWorkerService } from '../vendor/hachidori/extension/anki-worker.js';

async function mine({ proxy = true, audioFails = false } = {}) {
  const events: string[] = [];
  const filename = `hachidori_${'a'.repeat(64)}.mp3`;
  let fields: Record<string, string> = {};
  let initialAudio = '';
  let stored = false;
  const gateway = createAnkiGateway({
    readSubminerProxyUrl: async () => (proxy ? 'http://127.0.0.1:8766' : null),
    fetch: async (_url: string, options: RequestInit) => {
      const { action, params } = JSON.parse(String(options.body));
      events.push(action);
      let result: unknown;
      switch (action) {
        case 'multi':
          result = [
            { result: ['Default'], error: null },
            { result: ['Basic'], error: null },
            { result: ['Expression', 'ExpressionAudio'], error: null },
          ];
          break;
        case 'canAddNotes':
          result = [true];
          break;
        case 'canAddNotesWithErrorDetail':
          result = [{ canAdd: true, error: null }];
          break;
        case 'getMediaFilesNames':
          result = stored ? [filename] : [];
          break;
        case 'storeMediaFile':
          stored = true;
          result = filename;
          break;
        case 'addNote':
          fields = { ...params.note.fields };
          initialAudio = fields.ExpressionAudio ?? '';
          if (initialAudio) assert.ok(stored, 'audio must exist before the note references it');
          result = 123;
          break;
        case 'notesInfo':
          result = [
            {
              noteId: 123,
              fields: Object.fromEntries(
                Object.entries(fields).map(([key, value]) => [key, { value }]),
              ),
            },
          ];
          break;
        case 'updateNoteFields':
          Object.assign(fields, params.note.fields);
          result = null;
          break;
        default:
          throw new Error(`Unexpected action: ${action}`);
      }
      return Response.json({ result, error: null });
    },
  });
  const service = createAnkiWorkerService({
    gateway,
    readOptions: async () => ({
      anki: {
        url: 'http://127.0.0.1:8766',
        apiKey: '',
        templates: [
          {
            id: 'default',
            name: 'Default',
            deck: 'Default',
            model: 'Basic',
            tags: [],
            fields: {},
            duplicateScope: 'model',
            duplicateBehavior: 'prevent',
            captureScreenshot: false,
            fieldTemplates: {
              Expression: { value: '{expression}', overwriteMode: 'overwrite' },
              ExpressionAudio: { value: '{audio}', overwriteMode: 'overwrite' },
            },
          },
        ],
      },
      audioSources: [
        { id: 'test', enabled: true, type: 'custom', url: 'https://example.test/{term}' },
      ],
      mediaCapture: { enabled: false },
    }),
    readDictionaries: async () => [],
    engine: async () => ({ ready: true, loading: false, generation: 1 }),
    offscreen: async (message: { type: string; audio?: string }) => {
      if (message.type === 'hd_anki_audio') {
        events.push('pronunciation');
        if (audioFails) throw new Error('No pronunciation available');
        return { filename, data: 'YXVkaW8=' };
      }
      return { fields: { Expression: '猫', ExpressionAudio: message.audio ?? '' }, media: [] };
    },
    duplicateIndex: { source: async () => null, recordWrite: async () => {} },
  });
  const status = await service.status();
  assert.equal(status.available, true, status.error);
  const result = await service.submit({
    configKey: status.configKey,
    generation: 1,
    term: { expression: '猫', reading: 'ねこ' },
  });
  return { result, initialAudio, fields, events, filename };
}

test('SubMiner receives pronunciation in the initial Hachidori note, before enrichment starts', async () => {
  const value = await mine();
  assert.equal(value.result.state, 'added');
  assert.equal(value.initialAudio, `[sound:${value.filename}]`);
  assert.equal(value.events.filter((event) => event === 'pronunciation').length, 1);
});

test('direct Hachidori keeps deferred pronunciation', async () => {
  const value = await mine({ proxy: false });
  assert.equal(value.initialAudio, '');
  assert.equal(value.fields.ExpressionAudio, `[sound:${value.filename}]`);
});

test('unavailable pronunciation remains a warning without a late audio write', async () => {
  const value = await mine({ audioFails: true });
  assert.equal(value.result.state, 'added');
  assert.match(value.result.warnings.join(' '), /No pronunciation available/);
  assert.equal(value.initialAudio, '');
  assert.equal(value.events.filter((event) => event === 'pronunciation').length, 1);
  assert.equal(value.events.includes('updateNoteFields'), false);
});
