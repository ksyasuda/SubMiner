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

test('SessionEventPopover hides preview-unavailable fallback while note info is still loading', () => {
  const marker: SessionChartMarker = {
    key: 'card-177',
    kind: 'card',
    anchorTsMs: 9_000,
    eventTsMs: 9_000,
    noteIds: [177],
    cardsDelta: 1,
  };

  const markup = renderToStaticMarkup(
    <SessionEventPopover
      marker={marker}
      noteInfos={new Map()}
      loading
      pinned
      onTogglePinned={() => {}}
      onClose={() => {}}
      onOpenNote={() => {}}
    />,
  );

  assert.match(markup, /Loading Anki note info/);
  assert.doesNotMatch(markup, /Preview unavailable/);
});

test('SessionEventPopover keeps the loading state clean until note preview data arrives', () => {
  const marker: SessionChartMarker = {
    key: 'card-9001',
    kind: 'card',
    anchorTsMs: 9_001,
    eventTsMs: 9_001,
    noteIds: [1773808840964],
    cardsDelta: 1,
  };

  const markup = renderToStaticMarkup(
    <SessionEventPopover
      marker={marker}
      noteInfos={new Map()}
      loading={true}
      pinned={true}
      onTogglePinned={() => {}}
      onClose={() => {}}
      onOpenNote={() => {}}
    />,
  );

  assert.match(markup, /Loading Anki note info/);
  assert.doesNotMatch(markup, /Preview unavailable/);
});
