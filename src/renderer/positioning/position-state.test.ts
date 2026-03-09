import assert from 'node:assert/strict';
import test from 'node:test';

import { createInMemorySubtitlePositionController } from './position-state.js';

function withWindow<T>(windowValue: unknown, callback: () => T): T {
  const previousWindow = (globalThis as { window?: unknown }).window;
  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: windowValue,
  });

  try {
    return callback();
  } finally {
    Object.defineProperty(globalThis, 'window', {
      configurable: true,
      value: previousWindow,
    });
  }
}

function createContext(subtitleHeight: number) {
  return {
    dom: {
      subtitleContainer: {
        style: {
          position: '',
          left: '',
          top: '',
          right: '',
          transform: '',
          marginBottom: '',
        },
        offsetHeight: subtitleHeight,
      },
    },
    state: {
      currentYPercent: null,
      persistedSubtitlePosition: { yPercent: 10 },
    },
  };
}

test('subtitle position clamp keeps tall subtitles inside the overlay viewport', () => {
  withWindow(
    {
      innerHeight: 1000,
      electronAPI: {
        saveSubtitlePosition: () => {},
      },
    },
    () => {
      const ctx = createContext(300);
      const controller = createInMemorySubtitlePositionController(ctx as never);

      controller.applyYPercent(80);

      assert.equal(ctx.state.currentYPercent, 68.8);
      assert.equal(ctx.dom.subtitleContainer.style.marginBottom, '688px');
    },
  );
});

test('subtitle position clamp falls back to the minimum safe inset when subtitle is taller than the viewport', () => {
  withWindow(
    {
      innerHeight: 200,
      electronAPI: {
        saveSubtitlePosition: () => {},
      },
    },
    () => {
      const ctx = createContext(260);
      const controller = createInMemorySubtitlePositionController(ctx as never);

      controller.applyYPercent(80);

      assert.equal(ctx.state.currentYPercent, 6);
      assert.equal(ctx.dom.subtitleContainer.style.marginBottom, '12px');
    },
  );
});
