import assert from 'node:assert/strict';
import test from 'node:test';
import { createElement } from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { KanjiDetailPanel } from './KanjiDetailPanel';

test('KanjiDetailPanel uses the centered detail modal layout', () => {
  const markup = renderToStaticMarkup(
    createElement(KanjiDetailPanel, { kanjiId: 1, onClose: () => {} }),
  );

  assert.match(
    markup,
    /class="[^"]*fixed[^"]*inset-0[^"]*z-40[^"]*flex[^"]*items-center[^"]*justify-center[^"]*p-4/,
  );
  assert.match(
    markup,
    /class="[^"]*relative[^"]*flex[^"]*max-h-\[85vh\][^"]*w-full[^"]*max-w-2xl[^"]*flex-col/,
  );
  assert.doesNotMatch(markup, /class="[^"]*absolute[^"]*right-0[^"]*top-0[^"]*h-full/);
});
