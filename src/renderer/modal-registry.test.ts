import assert from 'node:assert/strict';
import test from 'node:test';

import { createModalRegistry, type ModalDescriptor } from './modal-registry';

test('modal registry derives open, subtitle-suppression, and active state from descriptor order', () => {
  const openIds = new Set(['animetosho', 'runtime-options']);
  const descriptors: ModalDescriptor<string>[] = [
    {
      id: 'jimaku',
      isOpen: () => openIds.has('jimaku'),
      close: () => {},
      suppressesSubtitles: true,
    },
    {
      id: 'animetosho',
      isOpen: () => openIds.has('animetosho'),
      close: () => {},
      suppressesSubtitles: false,
    },
    {
      id: 'runtime-options',
      isOpen: () => openIds.has('runtime-options'),
      close: () => {},
      suppressesSubtitles: true,
    },
  ];
  const registry = createModalRegistry(descriptors);

  assert.equal(registry.isAnyOpen(), true);
  assert.equal(registry.isAnySuppressingSubtitlesOpen(), true);
  assert.equal(registry.getActive(), 'animetosho');

  openIds.clear();
  assert.equal(registry.isAnyOpen(), false);
  assert.equal(registry.isAnySuppressingSubtitlesOpen(), false);
  assert.equal(registry.getActive(), null);
});

test('modal registry dismisses every open descriptor and skips closed descriptors', () => {
  const closed: string[] = [];
  const descriptors: ModalDescriptor<string>[] = ['closed', 'first-open', 'second-open'].map(
    (id) => ({
      id,
      isOpen: () => id !== 'closed',
      close: () => closed.push(id),
      suppressesSubtitles: false,
    }),
  );
  const registry = createModalRegistry(descriptors);

  registry.dismissOpen();

  assert.deepEqual(closed, ['first-open', 'second-open']);
});

test('modal registry can dismiss stale modals while preserving the requested one', () => {
  const openIds = new Set(['anime-browser', 'subtitle-sidebar']);
  const closed: string[] = [];
  const descriptors: ModalDescriptor<string>[] = ['anime-browser', 'subtitle-sidebar'].map(
    (id) => ({
      id,
      isOpen: () => openIds.has(id),
      close: () => {
        openIds.delete(id);
        closed.push(id);
      },
      suppressesSubtitles: true,
    }),
  );
  const registry = createModalRegistry(descriptors);

  registry.dismissOpenExcept('anime-browser');

  assert.deepEqual(closed, ['subtitle-sidebar']);
  assert.equal(openIds.has('anime-browser'), true);
});
