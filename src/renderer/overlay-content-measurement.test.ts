import assert from 'node:assert/strict';
import test from 'node:test';
import type { OverlayContentMeasurement, OverlayContentRect } from '../types';
import { createOverlayContentMeasurementReporter } from './overlay-content-measurement.js';

type FakeElement = {
  textContent: string;
  children: unknown[];
  getBoundingClientRect: () => DOMRect;
};

type DomKey =
  | 'subtitleRoot'
  | 'subtitleContainer'
  | 'secondarySubRoot'
  | 'secondarySubContainer'
  | 'subtitleSidebarContent'
  | 'overlayNotificationStack'
  | 'overlayNotificationHistory';

const ZERO: OverlayContentRect = { x: 0, y: 0, width: 0, height: 0 };

function el(
  rect: OverlayContentRect = ZERO,
  { text = '', children = 0 }: { text?: string; children?: number } = {},
): FakeElement {
  const domRect = { left: rect.x, top: rect.y, width: rect.width, height: rect.height } as DOMRect;
  return {
    textContent: text,
    children: Array.from({ length: children }, () => ({})),
    getBoundingClientRect: () => domRect,
  };
}

// Runs emitNow() once against a fake `window` and zero-rect empty DOM defaults.
// Returns the reported measurement, or undefined when nothing was reported.
function measure(
  dom: Partial<Record<DomKey, FakeElement>> = {},
  state: { subtitleSidebarModalOpen?: boolean; notificationHistoryOpen?: boolean } = {},
  overlayLayer = 'visible',
): Omit<OverlayContentMeasurement, 'measuredAtMs'> | undefined {
  const originalWindow = Object.getOwnPropertyDescriptor(globalThis, 'window');
  const reports: OverlayContentMeasurement[] = [];
  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    writable: true,
    value: {
      innerWidth: 1920,
      innerHeight: 1080,
      electronAPI: {
        reportOverlayContentBounds: (payload: OverlayContentMeasurement) => reports.push(payload),
      },
    },
  });

  try {
    const defaults: Record<DomKey, FakeElement> = {
      subtitleRoot: el(),
      subtitleContainer: el(),
      secondarySubRoot: el(),
      secondarySubContainer: el(),
      subtitleSidebarContent: el(),
      overlayNotificationStack: el(),
      overlayNotificationHistory: el(),
    };
    createOverlayContentMeasurementReporter({
      platform: { overlayLayer },
      state,
      dom: { ...defaults, ...dom },
    } as never).emitNow();
  } finally {
    if (originalWindow) {
      Object.defineProperty(globalThis, 'window', originalWindow);
    } else {
      delete (globalThis as { window?: unknown }).window;
    }
  }

  assert.ok(reports.length <= 1);
  const report = reports[0];
  if (!report) return undefined;
  assert.equal(typeof report.measuredAtMs, 'number');
  const { measuredAtMs: _measuredAtMs, ...rest } = report;
  return rest;
}

function expected(contentRect: OverlayContentRect | null, interactiveRects: OverlayContentRect[]) {
  return {
    layer: 'visible' as const,
    viewport: { width: 1920, height: 1080 },
    contentRect,
    interactiveRects,
  };
}

const PRIMARY = { x: 760, y: 890, width: 400, height: 92 };
const SECONDARY = { x: 700, y: 40, width: 520, height: 70 };
const SIDEBAR = { x: 1500, y: 60, width: 380, height: 900 };
const NOTIFICATIONS = { x: 1540, y: 16, width: 360, height: 220 };
const HISTORY = { x: 1400, y: 100, width: 480, height: 600 };

const cases: Array<{
  name: string;
  dom?: Partial<Record<DomKey, FakeElement>>;
  state?: { subtitleSidebarModalOpen?: boolean; notificationHistoryOpen?: boolean };
  layer?: string;
  report: ReturnType<typeof measure>;
}> = [
  {
    name: 'reports primary and secondary subtitle containers as separate rects and unions them',
    dom: {
      subtitleRoot: el({ x: 810, y: 910, width: 300, height: 48 }, { text: 'primary' }),
      subtitleContainer: el(PRIMARY),
      secondarySubRoot: el({ x: 850, y: 50, width: 220, height: 34 }, { text: 'English' }),
      secondarySubContainer: el(SECONDARY),
    },
    report: expected({ x: 700, y: 40, width: 520, height: 942 }, [PRIMARY, SECONDARY]),
  },
  {
    name: 'excludes a subtitle container whose root has only whitespace text',
    dom: {
      subtitleRoot: el(ZERO, { text: '  \n ' }),
      subtitleContainer: el(PRIMARY),
      secondarySubRoot: el(ZERO, { text: 'English' }),
      secondarySubContainer: el(SECONDARY),
    },
    report: expected(SECONDARY, [SECONDARY]),
  },
  {
    name: 'drops zero-area rects',
    dom: {
      subtitleRoot: el(ZERO, { text: 'primary' }),
      subtitleContainer: el({ x: 760, y: 890, width: 400, height: 0 }),
    },
    report: expected(null, []),
  },
  {
    name: 'drops non-finite rects',
    dom: {
      subtitleRoot: el(ZERO, { text: 'primary' }),
      subtitleContainer: el({ x: Number.NaN, y: 890, width: 400, height: 92 }),
      secondarySubRoot: el(ZERO, { text: 'English' }),
      secondarySubContainer: el({ x: 700, y: 40, width: Number.POSITIVE_INFINITY, height: 70 }),
    },
    report: expected(null, []),
  },
  {
    name: 'rounds rect coordinates to two decimals',
    dom: {
      subtitleRoot: el(ZERO, { text: 'primary' }),
      subtitleContainer: el({ x: 760.123, y: 890.456, width: 400.789, height: 92.001 }),
    },
    report: expected({ x: 760.12, y: 890.46, width: 400.79, height: 92 }, [
      { x: 760.12, y: 890.46, width: 400.79, height: 92 },
    ]),
  },
  {
    name: 'includes the open subtitle sidebar',
    dom: { subtitleSidebarContent: el(SIDEBAR) },
    state: { subtitleSidebarModalOpen: true },
    report: expected(SIDEBAR, [SIDEBAR]),
  },
  {
    name: 'ignores the sidebar when it is closed',
    dom: { subtitleSidebarContent: el(SIDEBAR) },
    state: { subtitleSidebarModalOpen: false },
    report: expected(null, []),
  },
  {
    name: 'includes the notification stack when it has cards',
    dom: { overlayNotificationStack: el(NOTIFICATIONS, { children: 2 }) },
    report: expected(NOTIFICATIONS, [NOTIFICATIONS]),
  },
  {
    name: 'ignores an empty notification stack',
    dom: { overlayNotificationStack: el(NOTIFICATIONS) },
    report: expected(null, []),
  },
  {
    name: 'includes the open notification history panel',
    dom: { overlayNotificationHistory: el(HISTORY) },
    state: { notificationHistoryOpen: true },
    report: expected(HISTORY, [HISTORY]),
  },
  {
    name: 'reports a null content rect when nothing is measurable',
    report: expected(null, []),
  },
  {
    name: 'emits nothing for a non-visible overlay layer',
    dom: {
      subtitleRoot: el(ZERO, { text: 'primary' }),
      subtitleContainer: el(PRIMARY),
    },
    layer: 'modal',
    report: undefined,
  },
];

for (const c of cases) {
  test(`overlay measurement ${c.name}`, () => {
    assert.deepEqual(measure(c.dom, c.state, c.layer), c.report);
  });
}
