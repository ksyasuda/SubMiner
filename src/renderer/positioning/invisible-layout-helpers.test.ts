import test from 'node:test';
import assert from 'node:assert/strict';

import type { MpvSubtitleRenderMetrics } from '../../types';
import {
  applyTypography,
  applyVerticalPosition,
  resolveBaselineCompensationPx,
} from './invisible-layout-helpers.js';

const METRICS: MpvSubtitleRenderMetrics = {
  subPos: 100,
  subFontSize: 38,
  subScale: 1,
  subMarginY: 34,
  subMarginX: 19,
  subFont: 'sans-serif',
  subSpacing: 0,
  subBold: false,
  subItalic: false,
  subBorderSize: 2.5,
  subShadowOffset: 0,
  subAssOverride: 'yes',
  subScaleByWindow: true,
  subUseMargins: true,
  osdHeight: 720,
  osdDimensions: null,
};

type TypographyTestContext = {
  dom: {
    subtitleRoot: { style: CSSStyleDeclaration };
    subtitleContainer: { style: CSSStyleDeclaration };
  };
  state: {
    currentInvisibleSubtitleLineCount: number;
    invisibleMeasuredDescentPx: number | null;
  };
  platform: {
    isMacOSPlatform: boolean;
  };
};

function withMockedComputedLineHeight(lineHeightPx: number, callback: () => void): void {
  const originalGetComputedStyle = (globalThis as { getComputedStyle?: unknown }).getComputedStyle;
  Object.defineProperty(globalThis, 'getComputedStyle', {
    configurable: true,
    value: () =>
      ({
        lineHeight: `${lineHeightPx}px`,
      }) as CSSStyleDeclaration,
  });
  try {
    callback();
  } finally {
    if (typeof originalGetComputedStyle === 'function') {
      Object.defineProperty(globalThis, 'getComputedStyle', {
        configurable: true,
        value: originalGetComputedStyle,
      });
    } else {
      Reflect.deleteProperty(globalThis, 'getComputedStyle');
    }
  }
}

function createStyle(initial: Record<string, string> = {}): CSSStyleDeclaration {
  const values: Record<string, string> = { ...initial };
  const target = {
    setProperty: (name: string, value: string) => {
      values[name] = value;
    },
    getPropertyValue: (name: string) => values[name] ?? '',
  } as unknown as CSSStyleDeclaration;

  return new Proxy(target, {
    get(obj, prop) {
      if (typeof prop === 'string') {
        if (prop in obj) return obj[prop as keyof CSSStyleDeclaration];
        return values[prop] ?? '';
      }
      return obj[prop as keyof CSSStyleDeclaration];
    },
    set(_obj, prop, value) {
      if (typeof prop === 'string') {
        values[prop] = String(value);
        return true;
      }
      return false;
    },
  });
}

function createContext(options: {
  isMacOSPlatform: boolean;
  lineCount: number;
  bottomPx?: number;
  topPx?: number;
}): TypographyTestContext {
  const subtitleRoot = { style: createStyle() };
  const subtitleContainer = {
    style: createStyle({
      bottom: typeof options.bottomPx === 'number' ? `${options.bottomPx}px` : '',
      top: typeof options.topPx === 'number' ? `${options.topPx}px` : '',
    }),
  };

  return {
    dom: { subtitleRoot, subtitleContainer },
    state: {
      currentInvisibleSubtitleLineCount: options.lineCount,
      invisibleMeasuredDescentPx: null,
    },
    platform: {
      isMacOSPlatform: options.isMacOSPlatform,
    },
  };
}

test('resolveBaselineCompensationPx uses measured descent when present', () => {
  const compensation = resolveBaselineCompensationPx(10, 2.5, 1);
  assert.equal(compensation, 16);
});

test('resolveBaselineCompensationPx falls back to border and shadow compensation when descent missing', () => {
  const compensation = resolveBaselineCompensationPx(null, 2.5, 1);
  assert.equal(compensation, 17.5);
});

test('applyTypography keeps macOS default letter spacing neutral when mpv spacing is zero', () => {
  const ctx = createContext({
    isMacOSPlatform: true,
    lineCount: 1,
    bottomPx: 120,
  });

  withMockedComputedLineHeight(34, () => {
    applyTypography(ctx as never, {
      metrics: { ...METRICS, subSpacing: 0 },
      pxPerScaledPixel: 1,
      effectiveFontSize: 34,
    });
  });

  assert.equal(ctx.dom.subtitleRoot.style.getPropertyValue('letter-spacing'), '0px');
});

test('applyTypography applies full mpv letter spacing scale on macOS', () => {
  const ctx = createContext({
    isMacOSPlatform: true,
    lineCount: 1,
    bottomPx: 120,
  });

  withMockedComputedLineHeight(34, () => {
    applyTypography(ctx as never, {
      metrics: { ...METRICS, subSpacing: 1.5 },
      pxPerScaledPixel: 2,
      effectiveFontSize: 34,
    });
  });

  assert.equal(ctx.dom.subtitleRoot.style.getPropertyValue('letter-spacing'), '3px');
});

test('applyTypography uses tighter macOS line-height for invisible multiline alignment', () => {
  const ctx = createContext({
    isMacOSPlatform: true,
    lineCount: 3,
    bottomPx: 120,
  });

  withMockedComputedLineHeight(34, () => {
    applyTypography(ctx as never, {
      metrics: METRICS,
      pxPerScaledPixel: 1,
      effectiveFontSize: 34,
    });
  });

  assert.equal(ctx.dom.subtitleRoot.style.getPropertyValue('line-height'), '0.96');
});

test('applyVerticalPosition uses subtitle position margin and baseline compensation', () => {
  const ctx = createContext({
    isMacOSPlatform: true,
    lineCount: 1,
  });

  applyVerticalPosition(ctx as never, {
    metrics: { ...METRICS, subPos: 90 },
    renderAreaHeight: 720,
    topInset: 0,
    bottomInset: 10,
    marginY: 34,
    borderPx: 2.5,
    shadowPx: 0,
    measuredDescentPx: null,
    vAlign: 0,
  });

  const bottom = parseFloat(ctx.dom.subtitleContainer.style.bottom);
  assert.ok(Number.isFinite(bottom));
  assert.ok(bottom > 90 && bottom < 105);
});

test('applyVerticalPosition uses measured descent consistently across line counts', () => {
  const single = createContext({
    isMacOSPlatform: true,
    lineCount: 1,
  });
  const dense = createContext({
    isMacOSPlatform: true,
    lineCount: 3,
  });

  applyVerticalPosition(single as never, {
    metrics: METRICS,
    renderAreaHeight: 720,
    topInset: 0,
    bottomInset: 0,
    marginY: 34,
    borderPx: 2.5,
    shadowPx: 0,
    measuredDescentPx: 12,
    vAlign: 0,
  });
  applyVerticalPosition(dense as never, {
    metrics: METRICS,
    renderAreaHeight: 720,
    topInset: 0,
    bottomInset: 0,
    marginY: 34,
    borderPx: 2.5,
    shadowPx: 0,
    measuredDescentPx: 12,
    vAlign: 0,
  });

  const singleBottom = parseFloat(single.dom.subtitleContainer.style.bottom);
  const denseBottom = parseFloat(dense.dom.subtitleContainer.style.bottom);
  assert.equal(singleBottom, denseBottom);
});
