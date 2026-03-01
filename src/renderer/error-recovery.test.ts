import test from 'node:test';
import assert from 'node:assert/strict';

import { createRendererRecoveryController } from './error-recovery.js';
import {
  YOMITAN_POPUP_IFRAME_SELECTOR,
  hasYomitanPopupIframe,
  isYomitanPopupIframe,
} from './yomitan-popup.js';
import { scrollActiveRuntimeOptionIntoView } from './modals/runtime-options.js';
import { resolvePlatformInfo } from './utils/platform.js';

test('handleError logs context and recovers overlay state', () => {
  const payloads: unknown[] = [];
  let dismissed = 0;
  let restored = 0;
  const shown: string[] = [];

  const controller = createRendererRecoveryController({
    dismissActiveUi: () => {
      dismissed += 1;
    },
    restoreOverlayInteraction: () => {
      restored += 1;
    },
    showToast: (message) => {
      shown.push(message);
    },
    getSnapshot: () => ({
      activeModal: 'jimaku',
      subtitlePreview: '字幕テキスト',
      secondarySubtitlePreview: 'secondary',
      isOverlayInteractive: true,
      isOverSubtitle: true,
      overlayLayer: 'visible',
    }),
    logError: (payload) => {
      payloads.push(payload);
    },
  });

  controller.handleError(new Error('renderer boom'), {
    source: 'callback',
    action: 'onSubtitle',
  });

  assert.equal(dismissed, 1);
  assert.equal(restored, 1);
  assert.equal(shown.length, 1);
  assert.match(shown[0]!, /recovered/i);
  assert.equal(payloads.length, 1);

  const payload = payloads[0] as {
    context: { action: string };
    error: { message: string; stack: string | null };
    snapshot: { activeModal: string | null; subtitlePreview: string };
  };
  assert.equal(payload.context.action, 'onSubtitle');
  assert.equal(payload.snapshot.activeModal, 'jimaku');
  assert.equal(payload.snapshot.subtitlePreview, '字幕テキスト');
  assert.equal(payload.error.message, 'renderer boom');
  assert.ok(
    typeof payload.error.stack === 'string' && payload.error.stack.includes('renderer boom'),
  );
});

test('handleError normalizes non-Error values', () => {
  const payloads: unknown[] = [];

  const controller = createRendererRecoveryController({
    dismissActiveUi: () => {},
    restoreOverlayInteraction: () => {},
    showToast: () => {},
    getSnapshot: () => ({
      activeModal: null,
      subtitlePreview: '',
      secondarySubtitlePreview: '',
      isOverlayInteractive: false,
      isOverSubtitle: false,
      overlayLayer: 'visible',
    }),
    logError: (payload) => {
      payloads.push(payload);
    },
  });

  controller.handleError({ code: 500, reason: 'timeout' }, { source: 'callback', action: 'modal' });

  const payload = payloads[0] as { error: { message: string; stack: string | null } };
  assert.equal(payload.error.message, JSON.stringify({ code: 500, reason: 'timeout' }));
  assert.equal(payload.error.stack, null);
});

test('nested recovery errors are ignored while current recovery is active', () => {
  const payloads: unknown[] = [];
  let restored = 0;

  let controllerRef: ReturnType<typeof createRendererRecoveryController> | null = null;

  const controller = createRendererRecoveryController({
    dismissActiveUi: () => {
      controllerRef?.handleError(new Error('nested'), { source: 'callback', action: 'nested' });
    },
    restoreOverlayInteraction: () => {
      restored += 1;
    },
    showToast: () => {},
    getSnapshot: () => ({
      activeModal: 'runtime-options',
      subtitlePreview: '',
      secondarySubtitlePreview: '',
      isOverlayInteractive: true,
      isOverSubtitle: false,
      overlayLayer: 'visible',
    }),
    logError: (payload) => {
      payloads.push(payload);
    },
  });
  controllerRef = controller;

  controller.handleError(new Error('outer'), { source: 'callback', action: 'outer' });

  assert.equal(payloads.length, 1);
  assert.equal(restored, 1);
});

test('resolvePlatformInfo prefers query layer over preload layer', () => {
  const previousWindow = (globalThis as { window?: unknown }).window;
  const previousNavigator = (globalThis as { navigator?: unknown }).navigator;

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      electronAPI: {
        getOverlayLayer: () => 'modal',
      },
      location: { search: '?layer=visible' },
    },
  });
  Object.defineProperty(globalThis, 'navigator', {
    configurable: true,
    value: {
      platform: 'MacIntel',
      userAgent: 'Mozilla/5.0 (Macintosh)',
    },
  });

  try {
    const info = resolvePlatformInfo();
    assert.equal(info.overlayLayer, 'visible');
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
    Object.defineProperty(globalThis, 'navigator', {
      configurable: true,
      value: previousNavigator,
    });
  }
});

test('resolvePlatformInfo ignores legacy secondary layer and falls back to visible', () => {
  const previousWindow = (globalThis as { window?: unknown }).window;
  const previousNavigator = (globalThis as { navigator?: unknown }).navigator;

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      electronAPI: {
        getOverlayLayer: () => 'secondary',
      },
      location: { search: '' },
    },
  });
  Object.defineProperty(globalThis, 'navigator', {
    configurable: true,
    value: {
      platform: 'MacIntel',
      userAgent: 'Mozilla/5.0 (Macintosh)',
    },
  });

  try {
    const info = resolvePlatformInfo();
    assert.equal(info.overlayLayer, 'visible');
    assert.equal(info.shouldToggleMouseIgnore, true);
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
    Object.defineProperty(globalThis, 'navigator', {
      configurable: true,
      value: previousNavigator,
    });
  }
});

test('resolvePlatformInfo supports modal layer and disables mouse-ignore toggles', () => {
  const previousWindow = (globalThis as { window?: unknown }).window;
  const previousNavigator = (globalThis as { navigator?: unknown }).navigator;

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      electronAPI: {
        getOverlayLayer: () => 'modal',
      },
      location: { search: '' },
    },
  });
  Object.defineProperty(globalThis, 'navigator', {
    configurable: true,
    value: {
      platform: 'MacIntel',
      userAgent: 'Mozilla/5.0 (Macintosh)',
    },
  });

  try {
    const info = resolvePlatformInfo();
    assert.equal(info.overlayLayer, 'modal');
    assert.equal(info.isModalLayer, true);
    assert.equal(info.shouldToggleMouseIgnore, false);
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
    Object.defineProperty(globalThis, 'navigator', {
      configurable: true,
      value: previousNavigator,
    });
  }
});

test('isYomitanPopupIframe matches modern popup class and legacy id prefix', () => {
  const createElement = (options: {
    tagName: string;
    id?: string;
    classNames?: string[];
  }): Element =>
    ({
      tagName: options.tagName,
      id: options.id ?? '',
      classList: {
        contains: (className: string) => (options.classNames ?? []).includes(className),
      },
    }) as unknown as Element;

  assert.equal(
    isYomitanPopupIframe(
      createElement({
        tagName: 'IFRAME',
        classNames: ['yomitan-popup'],
      }),
    ),
    true,
  );
  assert.equal(
    isYomitanPopupIframe(
      createElement({
        tagName: 'IFRAME',
        id: 'yomitan-popup-123',
      }),
    ),
    true,
  );
  assert.equal(
    isYomitanPopupIframe(
      createElement({
        tagName: 'IFRAME',
        id: 'something-else',
      }),
    ),
    false,
  );
});

test('hasYomitanPopupIframe queries for modern + legacy selector', () => {
  let selector = '';
  const root = {
    querySelector: (value: string) => {
      selector = value;
      return {};
    },
  } as unknown as ParentNode;

  assert.equal(hasYomitanPopupIframe(root), true);
  assert.equal(selector, YOMITAN_POPUP_IFRAME_SELECTOR);
});

test('scrollActiveRuntimeOptionIntoView scrolls active runtime option with nearest block', () => {
  const calls: Array<{ block?: ScrollLogicalPosition }> = [];
  const activeItem = {
    scrollIntoView: (options?: ScrollIntoViewOptions) => {
      calls.push({ block: options?.block });
    },
  };

  const list = {
    querySelector: (selector: string) => {
      assert.equal(selector, '.runtime-options-item.active');
      return activeItem as unknown as Element;
    },
  };

  scrollActiveRuntimeOptionIntoView(list);
  assert.deepEqual(calls, [{ block: 'nearest' }]);
});

test('scrollActiveRuntimeOptionIntoView no-ops without active option', () => {
  const list = {
    querySelector: () => null,
  };

  assert.doesNotThrow(() => {
    scrollActiveRuntimeOptionIntoView(list);
  });
});
