import assert from 'node:assert/strict';
import test from 'node:test';
import { renderToStaticMarkup } from 'react-dom/server';
import { SessionDetail } from '../components/sessions/SessionDetail';

test('SessionDetail omits the misleading new words metric', () => {
  const markup = renderToStaticMarkup(
    <SessionDetail
      session={{
        sessionId: 7,
        canonicalTitle: 'Episode 7',
        videoId: 7,
        animeId: null,
        animeTitle: null,
        startedAtMs: 0,
        endedAtMs: null,
        totalWatchedMs: 0,
        activeWatchedMs: 0,
        linesSeen: 12,
        wordsSeen: 24,
        tokensSeen: 24,
        cardsMined: 0,
        lookupCount: 0,
        lookupHits: 0,
        yomitanLookupCount: 0,
      }}
    />,
  );

  assert.match(markup, /Total words/);
  assert.doesNotMatch(markup, /New words/);
});
