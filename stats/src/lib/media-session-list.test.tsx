import assert from 'node:assert/strict';
import test from 'node:test';
import { renderToStaticMarkup } from 'react-dom/server';
import { MediaSessionList } from '../components/library/MediaSessionList';

test('MediaSessionList renders expandable session rows with delete affordance', () => {
  const markup = renderToStaticMarkup(
    <MediaSessionList
      sessions={[
        {
          sessionId: 7,
          canonicalTitle: 'Episode 7',
          videoId: 9,
          animeId: 3,
          animeTitle: 'Anime',
          startedAtMs: 0,
          endedAtMs: null,
          totalWatchedMs: 1_000,
          activeWatchedMs: 900,
          linesSeen: 12,
          tokensSeen: 24,
          cardsMined: 2,
          lookupCount: 3,
          lookupHits: 2,
          yomitanLookupCount: 1,
          knownWordsSeen: 6,
          knownWordRate: 25,
        },
      ]}
      onDeleteSession={() => {}}
      initialExpandedSessionId={7}
    />,
  );

  assert.match(markup, /Session History/);
  assert.match(markup, /aria-expanded="true"/);
  assert.match(markup, /Delete session Episode 7/);
  assert.match(markup, /tokens/);
  assert.match(markup, /No token data for this session/);
});
