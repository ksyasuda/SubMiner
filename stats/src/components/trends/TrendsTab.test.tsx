import assert from 'node:assert/strict';
import test from 'node:test';
import { renderToStaticMarkup } from 'react-dom/server';
import { AnimeVisibilityFilter } from './TrendsTab';

test('AnimeVisibilityFilter uses title visibility wording', () => {
  const markup = renderToStaticMarkup(
    <AnimeVisibilityFilter
      animeTitles={['KonoSuba']}
      hiddenAnime={new Set()}
      maxTitles={null}
      onShowAll={() => {}}
      onHideAll={() => {}}
      onToggleAnime={() => {}}
      onMaxTitlesChange={() => {}}
    />,
  );

  assert.match(markup, /Title Visibility/);
  assert.doesNotMatch(markup, /Anime Visibility/);
});

test('AnimeVisibilityFilter offers a per-chart title limit selector', () => {
  const markup = renderToStaticMarkup(
    <AnimeVisibilityFilter
      animeTitles={['KonoSuba']}
      hiddenAnime={new Set()}
      maxTitles={7}
      onShowAll={() => {}}
      onHideAll={() => {}}
      onToggleAnime={() => {}}
      onMaxTitlesChange={() => {}}
    />,
  );

  assert.match(markup, /per chart/);
  assert.match(markup, /<option value="all">All<\/option>/);
  assert.match(markup, /<option value="7" selected="">/);
});

test('TrendsTab source labels words per minute without reading speed wording', async () => {
  const source = await Bun.file(new URL('./TrendsTab.tsx', import.meta.url)).text();

  assert.match(source, /title="Words \/ Min"/);
  assert.doesNotMatch(source, /Reading Speed/);
});
