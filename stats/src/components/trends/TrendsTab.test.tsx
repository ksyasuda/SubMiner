import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';
import { renderToStaticMarkup } from 'react-dom/server';
import { AnimeVisibilityFilter } from './TrendsTab';

test('AnimeVisibilityFilter uses title visibility wording', () => {
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

  assert.match(markup, /Title Visibility/);
  assert.doesNotMatch(markup, /Anime Visibility/);
});

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

test('AnimeVisibilityFilter offers top vs most-recent ranking modes', () => {
  const markup = renderToStaticMarkup(
    <AnimeVisibilityFilter
      animeTitles={['KonoSuba']}
      hiddenAnime={new Set()}
      maxTitles={7}
      maxTitlesMode="recent"
      onShowAll={() => {}}
      onHideAll={() => {}}
      onToggleAnime={() => {}}
      onMaxTitlesChange={() => {}}
      onMaxTitlesModeChange={() => {}}
    />,
  );

  assert.match(markup, /<option value="total">top<\/option>/);
  assert.match(markup, /<option value="recent" selected="">most recent<\/option>/);
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

test('TrendsTab source labels words per minute without reading speed wording', async () => {
  const source = await readFile(new URL('./TrendsTab.tsx', import.meta.url), 'utf8');

  assert.match(source, /title="Words \/ Min"/);
  assert.doesNotMatch(source, /Reading Speed/);
});
