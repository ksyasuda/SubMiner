import assert from 'node:assert/strict';
import test from 'node:test';

import { createOverlayNotificationsController } from './overlay-notifications.js';

function createClassList(initialTokens: string[] = []) {
  const tokens = new Set(initialTokens);
  return {
    add: (...entries: string[]) => {
      for (const entry of entries) tokens.add(entry);
    },
    remove: (...entries: string[]) => {
      for (const entry of entries) tokens.delete(entry);
    },
    contains: (entry: string) => tokens.has(entry),
  };
}

test('overlay notifications show loading state with spinner and no auto-hide timer', () => {
  const toast = {
    classList: createClassList(['hidden']),
    dataset: {} as Record<string, string>,
  };
  const title = { textContent: '' };
  const message = { textContent: '' };
  const spinner = { classList: createClassList(['hidden']) };
  let scheduled = false;

  const controller = createOverlayNotificationsController(
    {
      overlayNotificationToast: toast,
      overlayNotificationTitle: title,
      overlayNotificationMessage: message,
      overlayNotificationSpinner: spinner,
    } as never,
    {
      setTimeout: () => {
        scheduled = true;
        return 1 as never;
      },
      clearTimeout: () => {},
    },
  );

  controller.show({
    kind: 'loading',
    title: 'SubMiner',
    message: 'Loading subtitle annotations',
  });

  assert.equal(toast.classList.contains('hidden'), false);
  assert.equal(spinner.classList.contains('hidden'), false);
  assert.equal(title.textContent, 'SubMiner');
  assert.equal(message.textContent, 'Loading subtitle annotations');
  assert.equal(toast.dataset.kind, 'loading');
  assert.equal(scheduled, false);
});

test('overlay notifications auto-hide non-loading messages and clear loading styling', () => {
  let nextTimerId = 1;
  const scheduled = new Map<number, () => void>();
  const toast = {
    classList: createClassList(['hidden']),
    dataset: {} as Record<string, string>,
  };
  const title = { textContent: '' };
  const message = { textContent: '' };
  const spinner = { classList: createClassList(['hidden']) };

  const controller = createOverlayNotificationsController(
    {
      overlayNotificationToast: toast,
      overlayNotificationTitle: title,
      overlayNotificationMessage: message,
      overlayNotificationSpinner: spinner,
    } as never,
    {
      durationMs: 1200,
      setTimeout: (callback: () => void) => {
        const id = nextTimerId++;
        scheduled.set(id, callback);
        return id as never;
      },
      clearTimeout: (id) => {
        scheduled.delete(id as never as number);
      },
    },
  );

  controller.show({
    kind: 'loading',
    title: 'SubMiner',
    message: 'Loading subtitle annotations',
  });
  controller.show({
    kind: 'success',
    title: 'SubMiner',
    message: 'Subtitle annotations loaded',
  });

  assert.equal(spinner.classList.contains('hidden'), true);
  assert.equal(toast.dataset.kind, 'success');
  assert.equal(message.textContent, 'Subtitle annotations loaded');
  assert.equal(scheduled.size, 1);

  const [hide] = scheduled.values();
  hide?.();

  assert.equal(toast.classList.contains('hidden'), true);
  assert.equal(title.textContent, '');
  assert.equal(message.textContent, '');
});
