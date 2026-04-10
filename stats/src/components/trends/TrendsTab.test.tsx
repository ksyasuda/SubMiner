import assert from 'node:assert/strict';
import test from 'node:test';
import { renderToStaticMarkup } from 'react-dom/server';
import { AnimeVisibilityFilter } from './TrendsTab';

test('AnimeVisibilityFilter uses title visibility wording', () => {
  const markup = renderToStaticMarkup(
    <AnimeVisibilityFilter
      animeTitles={['KonoSuba']}
      hiddenAnime={new Set()}
      onShowAll={() => {}}
      onHideAll={() => {}}
      onToggleAnime={() => {}}
    />,
  );

  assert.match(markup, /Title Visibility/);
  assert.doesNotMatch(markup, /Anime Visibility/);
});
