import assert from 'node:assert/strict';
import test from 'node:test';
import { renderToStaticMarkup } from 'react-dom/server';
import { TabBar } from '../components/layout/TabBar';
import { EpisodeList } from '../components/anime/EpisodeList';

test('TabBar renders Library instead of Anime for the media library tab', () => {
  const markup = renderToStaticMarkup(<TabBar activeTab="overview" onTabChange={() => {}} />);

  assert.doesNotMatch(markup, />Anime</);
  assert.match(markup, />Overview</);
  assert.match(markup, />Library</);
  assert.match(markup, />Search</);
});

test('EpisodeList renders explicit episode detail button alongside quick peek row', () => {
  const markup = renderToStaticMarkup(
    <EpisodeList
      episodes={[
        {
          videoId: 9,
          episode: 9,
          season: 1,
          durationMs: 1,
          endedMediaMs: null,
          watched: 0,
          canonicalTitle: 'Episode 9',
          totalSessions: 1,
          totalActiveMs: 1,
          totalCards: 1,
          totalTokensSeen: 350,
          totalYomitanLookupCount: 7,
          lastWatchedMs: 0,
        },
      ]}
      onOpenDetail={() => {}}
    />,
  );

  assert.match(markup, />Details</);
  assert.match(markup, /Episode 9/);
});
