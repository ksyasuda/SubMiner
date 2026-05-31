import assert from 'node:assert/strict';
import test from 'node:test';
import { createOverlayContentMeasurementReporter } from './overlay-content-measurement.js';

function makeElement(textContent: string, rect: DOMRect): HTMLElement {
  return {
    textContent,
    getBoundingClientRect: () => rect,
  } as unknown as HTMLElement;
}

test('overlay measurement reports primary and secondary subtitle bars as separate interactive rects', () => {
  const originalWindow = Object.getOwnPropertyDescriptor(globalThis, 'window');
  const reports: unknown[] = [];

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    writable: true,
    value: {
      innerWidth: 1920,
      innerHeight: 1080,
      electronAPI: {
        reportOverlayContentBounds: (payload: unknown) => {
          reports.push(payload);
        },
      },
    },
  });

  try {
    const reporter = createOverlayContentMeasurementReporter({
      platform: { overlayLayer: 'visible' },
      dom: {
        subtitleRoot: makeElement('primary', {
          left: 810,
          top: 910,
          width: 300,
          height: 48,
        } as DOMRect),
        subtitleContainer: makeElement('primary', {
          left: 760,
          top: 890,
          width: 400,
          height: 92,
        } as DOMRect),
        secondarySubRoot: makeElement('English', {
          left: 850,
          top: 50,
          width: 220,
          height: 34,
        } as DOMRect),
        secondarySubContainer: makeElement('English', {
          left: 700,
          top: 40,
          width: 520,
          height: 70,
        } as DOMRect),
      },
    } as never);

    reporter.emitNow();

    const measuredAtMs = (reports[0] as { measuredAtMs?: unknown } | undefined)?.measuredAtMs;
    if (typeof measuredAtMs !== 'number') {
      assert.fail('Expected report timestamp.');
    }

    assert.deepEqual(reports, [
      {
        layer: 'visible',
        measuredAtMs,
        viewport: { width: 1920, height: 1080 },
        contentRect: { x: 700, y: 40, width: 520, height: 942 },
        interactiveRects: [
          { x: 760, y: 890, width: 400, height: 92 },
          { x: 700, y: 40, width: 520, height: 70 },
        ],
      },
    ]);
  } finally {
    if (originalWindow) {
      Object.defineProperty(globalThis, 'window', originalWindow);
    } else {
      delete (globalThis as { window?: unknown }).window;
    }
  }
});

test('overlay measurement includes open subtitle sidebar bounds as an interactive rect', () => {
  const originalWindow = Object.getOwnPropertyDescriptor(globalThis, 'window');
  const reports: unknown[] = [];

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    writable: true,
    value: {
      innerWidth: 1920,
      innerHeight: 1080,
      electronAPI: {
        reportOverlayContentBounds: (payload: unknown) => {
          reports.push(payload);
        },
      },
    },
  });

  try {
    const reporter = createOverlayContentMeasurementReporter({
      platform: { overlayLayer: 'visible' },
      state: { subtitleSidebarModalOpen: true },
      dom: {
        subtitleRoot: makeElement('', {
          left: 0,
          top: 0,
          width: 0,
          height: 0,
        } as DOMRect),
        subtitleContainer: makeElement('', {
          left: 0,
          top: 0,
          width: 0,
          height: 0,
        } as DOMRect),
        secondarySubRoot: makeElement('', {
          left: 0,
          top: 0,
          width: 0,
          height: 0,
        } as DOMRect),
        secondarySubContainer: makeElement('', {
          left: 0,
          top: 0,
          width: 0,
          height: 0,
        } as DOMRect),
        subtitleSidebarContent: makeElement('sidebar', {
          left: 1500,
          top: 60,
          width: 380,
          height: 900,
        } as DOMRect),
      },
    } as never);

    reporter.emitNow();

    const measuredAtMs = (reports[0] as { measuredAtMs?: unknown } | undefined)?.measuredAtMs;
    if (typeof measuredAtMs !== 'number') {
      assert.fail('Expected report timestamp.');
    }

    assert.deepEqual(reports, [
      {
        layer: 'visible',
        measuredAtMs,
        viewport: { width: 1920, height: 1080 },
        contentRect: { x: 1500, y: 60, width: 380, height: 900 },
        interactiveRects: [{ x: 1500, y: 60, width: 380, height: 900 }],
      },
    ]);
  } finally {
    if (originalWindow) {
      Object.defineProperty(globalThis, 'window', originalWindow);
    } else {
      delete (globalThis as { window?: unknown }).window;
    }
  }
});
