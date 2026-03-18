import assert from 'node:assert/strict';
import test from 'node:test';
import { renderToStaticMarkup } from 'react-dom/server';
import type { SessionChartMarker } from '../../lib/session-events';
import { SessionEventPopover } from './SessionEventPopover';

test('SessionEventPopover renders formatted card-mine details with fetched note info', () => {
  const marker: SessionChartMarker = {
    key: 'card-6000',
    kind: 'card',
    anchorTsMs: 6_000,
    eventTsMs: 6_000,
    noteIds: [11, 22],
    cardsDelta: 2,
  };

  const markup = renderToStaticMarkup(
    <SessionEventPopover
      marker={marker}
      noteInfos={
        new Map([
          [11, { noteId: 11, expression: '冒険者', context: '駆け出しの冒険者だ', meaning: null }],
          [22, { noteId: 22, expression: '呪い', context: null, meaning: 'curse' }],
        ])
      }
      loading={false}
      pinned={false}
      onTogglePinned={() => {}}
      onClose={() => {}}
      onOpenNote={() => {}}
    />,
  );

  assert.match(markup, /Card mined/);
  assert.match(markup, /\+2 cards/);
  assert.match(markup, /冒険者/);
  assert.match(markup, /呪い/);
  assert.match(markup, /駆け出しの冒険者だ/);
  assert.match(markup, /curse/);
  assert.match(markup, /Pin/);
  assert.match(markup, /Open in Anki/);
});

test('SessionEventPopover renders seek metadata compactly', () => {
  const marker: SessionChartMarker = {
    key: 'seek-3000',
    kind: 'seek',
    anchorTsMs: 3_000,
    eventTsMs: 3_000,
    direction: 'backward',
    fromMs: 5_000,
    toMs: 1_500,
  };

  const markup = renderToStaticMarkup(
    <SessionEventPopover
      marker={marker}
      noteInfos={new Map()}
      loading={false}
      pinned={false}
      onTogglePinned={() => {}}
      onClose={() => {}}
      onOpenNote={() => {}}
    />,
  );

  assert.match(markup, /Seek backward/);
  assert.match(markup, /5\.0s/);
  assert.match(markup, /1\.5s/);
  assert.match(markup, /3\.5s/);
});

test('SessionEventPopover renders a cleaner fallback when AnkiConnect provides no preview fields', () => {
  const marker: SessionChartMarker = {
    key: 'card-9000',
    kind: 'card',
    anchorTsMs: 9_000,
    eventTsMs: 9_000,
    noteIds: [91],
    cardsDelta: 1,
  };

  const markup = renderToStaticMarkup(
    <SessionEventPopover
      marker={marker}
      noteInfos={new Map()}
      loading={false}
      pinned={true}
      onTogglePinned={() => {}}
      onClose={() => {}}
      onOpenNote={() => {}}
    />,
  );

  assert.match(markup, /Pinned/);
  assert.match(markup, /Preview unavailable from AnkiConnect/);
  assert.doesNotMatch(markup, /No readable note fields returned/);
});
