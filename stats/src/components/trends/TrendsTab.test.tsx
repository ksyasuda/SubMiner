import assert from 'node:assert/strict';
import test from 'node:test';
import { renderToStaticMarkup } from 'react-dom/server';
import { AnimeVisibilityFilter } from './TrendsTab';

test('AnimeVisibilityFilter offers a per-chart title limit selector', () => {
  const markup = renderToStaticMarkup(
    <AnimeVisibilityFilter
      animeTitles={['KonoSuba']}
      hiddenAnime={new Set()}
      maxTitles={7}
      maxTitlesMode="total"
      onShowAll={() => {}}
      onHideAll={() => {}}
      onToggleAnime={() => {}}
      onMaxTitlesChange={() => {}}
      onMaxTitlesModeChange={() => {}}
    />,
  );

  assert.match(markup, /per chart/);
  assert.match(markup, /<option value="all">All<\/option>/);
  assert.match(markup, /<option value="7" selected="">/);
});

test('AnimeVisibilityFilter keeps the ranking mode selectable even when showing all titles', () => {
  const markup = renderToStaticMarkup(
    <AnimeVisibilityFilter
      animeTitles={['KonoSuba']}
      hiddenAnime={new Set()}
      maxTitles={null}
      maxTitlesMode="total"
      onShowAll={() => {}}
      onHideAll={() => {}}
      onToggleAnime={() => {}}
      onMaxTitlesChange={() => {}}
      onMaxTitlesModeChange={() => {}}
    />,
  );

  assert.match(markup, /aria-label="Title ranking mode"/);
  assert.doesNotMatch(markup, /aria-label="Title ranking mode"[^>]*disabled/);
});
