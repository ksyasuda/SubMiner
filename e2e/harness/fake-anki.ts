import { once } from 'node:events';
import http from 'node:http';
import type { AddressInfo } from 'node:net';

export type AnkiRequest = { action: string; params: Record<string, unknown> };

export type FakeAnki = {
  url: string;
  /** Every AnkiConnect action received so far, with `multi` batches flattened. */
  requests: AnkiRequest[];
  stop: () => Promise<void>;
};

export const FAKE_ANKI_DECK = 'E2E';
export const FAKE_ANKI_MODEL = 'Lapis';

const MODEL_FIELDS = [
  'Expression',
  'ExpressionFurigana',
  'ExpressionReading',
  'ExpressionAudio',
  'SelectionText',
  'MainDefinition',
  'Sentence',
  'SentenceFurigana',
  'SentenceAudio',
  'Picture',
  'Glossary',
  'IsWordAndSentenceCard',
  'IsClickCard',
  'IsSentenceCard',
  'IsAudioCard',
  'MiscInfo',
];

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

// A recording AnkiConnect stand-in: an empty collection with one deck and one
// Lapis-shaped note type. Scenarios assert on `requests`, not on stored notes.
export async function startFakeAnki(): Promise<FakeAnki> {
  const requests: AnkiRequest[] = [];
  let nextNoteId = 1000;

  const handle = (body: unknown): unknown => {
    if (!isRecord(body) || typeof body.action !== 'string') return null;
    const params = isRecord(body.params) ? body.params : {};
    if (body.action === 'multi') {
      return Array.isArray(params.actions) ? params.actions.map(handle) : [];
    }
    requests.push({ action: body.action, params });
    switch (body.action) {
      case 'version':
        return 6;
      case 'deckNames':
        return [FAKE_ANKI_DECK];
      case 'modelNames':
        return [FAKE_ANKI_MODEL];
      case 'modelFieldNames':
        return MODEL_FIELDS;
      case 'addNote':
        return nextNoteId++;
      case 'findNotes':
      case 'notesInfo':
        return [];
      case 'storeMediaFile':
        return params.filename ?? null;
      default:
        return null;
    }
  };

  const server = http.createServer((request, response) => {
    response.setHeader('content-type', 'application/json');
    // GET exposes the request log so the session CLI can read it from outside.
    if (request.method === 'GET') {
      response.end(JSON.stringify(requests));
      return;
    }
    const chunks: Buffer[] = [];
    request.on('data', (chunk: Buffer) => chunks.push(chunk));
    request.on('end', () => {
      let body: unknown = null;
      try {
        body = JSON.parse(Buffer.concat(chunks).toString('utf8'));
      } catch {
        // Malformed bodies get the same null result AnkiConnect gives unknown actions.
      }
      response.end(JSON.stringify({ result: handle(body), error: null }));
    });
  });
  server.listen(0, '127.0.0.1');
  await once(server, 'listening');
  const { port } = server.address() as AddressInfo;

  return {
    url: `http://127.0.0.1:${port}`,
    requests,
    stop: async () => {
      server.closeAllConnections();
      server.close();
      await once(server, 'close');
    },
  };
}
