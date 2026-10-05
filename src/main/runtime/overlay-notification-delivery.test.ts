import assert from 'node:assert/strict';
import test from 'node:test';
import { createOverlayNotificationDelivery } from './overlay-notification-delivery';

test('overlay notification delivery queues until an overlay window is ready', () => {
  const sent: string[] = [];
  let ready = false;
  const delivery = createOverlayNotificationDelivery({
    hasReadyOverlayWindow: () => ready,
    send: (payload) => sent.push(`${payload.id ?? ''}:${'body' in payload ? payload.body : ''}`),
  });

  delivery.send({ id: 'startup-tokenization', title: 'Subtitle tokenization', body: 'Loading' });
  delivery.send({ id: 'character-dictionary-auto-sync', title: 'Dictionary', body: 'Building' });

  assert.equal(delivery.getQueuedCount(), 2);
  assert.deepEqual(sent, []);

  ready = true;
  delivery.flush();

  assert.equal(delivery.getQueuedCount(), 0);
  assert.deepEqual(sent, [
    'startup-tokenization:Loading',
    'character-dictionary-auto-sync:Building',
  ]);
});

test('overlay notification delivery upserts queued progress by notification id', () => {
  const sent: string[] = [];
  let ready = false;
  const delivery = createOverlayNotificationDelivery({
    hasReadyOverlayWindow: () => ready,
    send: (payload) => sent.push(`${payload.id ?? ''}:${'body' in payload ? payload.body : ''}`),
  });

  delivery.send({ id: 'startup-subtitle-annotations', title: 'Subtitle annotations', body: '|' });
  delivery.send({ id: 'startup-subtitle-annotations', title: 'Subtitle annotations', body: '/' });
  delivery.send({ id: 'startup-tokenization', title: 'Subtitle tokenization', body: 'Ready' });

  ready = true;
  delivery.flush();

  assert.deepEqual(sent, ['startup-subtitle-annotations:/', 'startup-tokenization:Ready']);
});

test('overlay notification delivery preserves queued events with distinct history ids', () => {
  const sent: string[] = [];
  let ready = false;
  const delivery = createOverlayNotificationDelivery({
    hasReadyOverlayWindow: () => ready,
    send: (payload) =>
      sent.push(
        `${payload.id ?? ''}:${'historyId' in payload ? payload.historyId : ''}:${'body' in payload ? payload.body : ''}`,
      ),
  });

  delivery.send({
    id: 'character-dictionary-auto-sync',
    historyId: 'character-dictionary-auto-sync-checking',
    title: 'Character dictionary',
    body: 'Checking character dictionary...',
    persistent: true,
  });
  delivery.send({
    id: 'character-dictionary-auto-sync',
    historyId: 'character-dictionary-auto-sync-building',
    title: 'Character dictionary',
    body: 'Building character dictionary...',
    persistent: true,
  });

  ready = true;
  delivery.flush();

  assert.deepEqual(sent, [
    'character-dictionary-auto-sync:character-dictionary-auto-sync-checking:Checking character dictionary...',
    'character-dictionary-auto-sync:character-dictionary-auto-sync-building:Building character dictionary...',
  ]);
});

test('overlay notification delivery preserves queued startup progress before terminal update', () => {
  const sent: string[] = [];
  const scheduled: Array<() => void> = [];
  let ready = false;
  const delivery = createOverlayNotificationDelivery({
    hasReadyOverlayWindow: () => ready,
    send: (payload) =>
      sent.push(
        `${payload.id ?? ''}:${'body' in payload ? payload.body : ''}:${'persistent' in payload && payload.persistent ? 'pin' : 'auto'}`,
      ),
    scheduleFlushRetry: (callback) => {
      scheduled.push(callback);
    },
  });

  delivery.send({
    id: 'startup-tokenization',
    title: 'Subtitle tokenization',
    body: 'Loading subtitle tokenization...',
    variant: 'progress',
    persistent: true,
  });
  delivery.send({
    id: 'startup-tokenization',
    title: 'Subtitle tokenization',
    body: 'Subtitle tokenization ready',
    variant: 'success',
    persistent: false,
  });

  ready = true;
  delivery.flush();
  scheduled.shift()?.();

  assert.deepEqual(sent, [
    'startup-tokenization:Loading subtitle tokenization...:pin',
    'startup-tokenization:Subtitle tokenization ready:auto',
  ]);
});

test('overlay notification delivery defers terminal update after first queued progress paint', () => {
  const sent: string[] = [];
  const scheduled: Array<() => void> = [];
  const delays: number[] = [];
  let ready = false;
  const delivery = createOverlayNotificationDelivery({
    hasReadyOverlayWindow: () => ready,
    send: (payload) =>
      sent.push(
        `${payload.id ?? ''}:${'body' in payload ? payload.body : ''}:${'persistent' in payload && payload.persistent ? 'pin' : 'auto'}`,
      ),
    scheduleFlushRetry: (callback, delayMs) => {
      scheduled.push(callback);
      delays.push(delayMs);
    },
    terminalUpdateDelayMs: 750,
  });

  delivery.send({
    id: 'startup-subtitle-annotations',
    title: 'Subtitle annotations',
    body: 'Loading subtitle annotations |',
    variant: 'progress',
    persistent: true,
  });
  delivery.send({
    id: 'startup-subtitle-annotations',
    title: 'Subtitle annotations',
    body: 'Subtitle annotations loaded',
    variant: 'success',
    persistent: false,
  });

  ready = true;
  delivery.flush();

  assert.deepEqual(sent, ['startup-subtitle-annotations:Loading subtitle annotations |:pin']);
  assert.equal(delivery.getQueuedCount(), 1);
  assert.deepEqual(delays, [750]);

  scheduled.shift()?.();

  assert.equal(delivery.getQueuedCount(), 0);
  assert.deepEqual(sent, [
    'startup-subtitle-annotations:Loading subtitle annotations |:pin',
    'startup-subtitle-annotations:Subtitle annotations loaded:auto',
  ]);
});

test('overlay notification delivery retries flush when lifecycle fires before window readiness settles', () => {
  const sent: string[] = [];
  const scheduled: Array<() => void> = [];
  let ready = false;
  const delivery = createOverlayNotificationDelivery({
    hasReadyOverlayWindow: () => ready,
    send: (payload) => sent.push(`${payload.id ?? ''}:${'body' in payload ? payload.body : ''}`),
    scheduleFlushRetry: (callback) => {
      scheduled.push(callback);
    },
  });

  delivery.send({ id: 'startup-tokenization', title: 'Subtitle tokenization', body: 'Loading' });
  delivery.flush();

  assert.equal(delivery.getQueuedCount(), 1);
  assert.equal(scheduled.length, 1);
  assert.deepEqual(sent, []);

  ready = true;
  scheduled.shift()?.();

  assert.equal(delivery.getQueuedCount(), 0);
  assert.deepEqual(sent, ['startup-tokenization:Loading']);
});

test('overlay notification delivery drops queued notification when dismissed before flush', () => {
  const sent: string[] = [];
  let ready = false;
  const delivery = createOverlayNotificationDelivery({
    hasReadyOverlayWindow: () => ready,
    send: (payload) =>
      sent.push('dismiss' in payload ? `dismiss:${payload.id}` : `show:${payload.id ?? ''}`),
  });

  delivery.send({ id: 'overlay-loading-status', title: 'SubMiner', body: 'Overlay loading' });
  delivery.send({ id: 'overlay-loading-status', dismiss: true });

  ready = true;
  delivery.flush();

  assert.deepEqual(sent, []);
});

test('overlay notification delivery removes queued notification when dismissed at readiness', () => {
  const sent: string[] = [];
  let ready = false;
  const delivery = createOverlayNotificationDelivery({
    hasReadyOverlayWindow: () => ready,
    send: (payload) =>
      sent.push('dismiss' in payload ? `dismiss:${payload.id}` : `show:${payload.id ?? ''}`),
  });

  delivery.send({ id: 'overlay-loading-status', title: 'SubMiner', body: 'Overlay loading' });

  ready = true;
  delivery.send({ id: 'overlay-loading-status', dismiss: true });
  delivery.flush();

  assert.deepEqual(sent, ['dismiss:overlay-loading-status']);
});
