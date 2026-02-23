import type { OverlayContentMeasurement, OverlayContentRect } from '../types';
import type { RendererContext } from './context';

const MEASUREMENT_DEBOUNCE_MS = 80;

function isMeasurableOverlayLayer(layer: string): layer is 'visible' | 'invisible' {
  return layer === 'visible' || layer === 'invisible';
}

function round2(value: number): number {
  return Math.round(value * 100) / 100;
}

function toMeasuredRect(rect: DOMRect): OverlayContentRect | null {
  if (!Number.isFinite(rect.left) || !Number.isFinite(rect.top)) {
    return null;
  }
  if (!Number.isFinite(rect.width) || !Number.isFinite(rect.height)) {
    return null;
  }

  const width = Math.max(0, rect.width);
  const height = Math.max(0, rect.height);

  return {
    x: round2(rect.left),
    y: round2(rect.top),
    width: round2(width),
    height: round2(height),
  };
}

function unionRects(a: OverlayContentRect, b: OverlayContentRect): OverlayContentRect {
  const left = Math.min(a.x, b.x);
  const top = Math.min(a.y, b.y);
  const right = Math.max(a.x + a.width, b.x + b.width);
  const bottom = Math.max(a.y + a.height, b.y + b.height);
  return {
    x: round2(left),
    y: round2(top),
    width: round2(Math.max(0, right - left)),
    height: round2(Math.max(0, bottom - top)),
  };
}

function hasVisibleTextContent(element: HTMLElement): boolean {
  return Boolean(element.textContent && element.textContent.trim().length > 0);
}

function collectContentRect(ctx: RendererContext): OverlayContentRect | null {
  let combinedRect: OverlayContentRect | null = null;

  const subtitleHasContent = hasVisibleTextContent(ctx.dom.subtitleRoot);
  if (subtitleHasContent) {
    const subtitleRect = toMeasuredRect(ctx.dom.subtitleRoot.getBoundingClientRect());
    if (subtitleRect) {
      combinedRect = subtitleRect;
    }
  }

  const secondaryHasContent = hasVisibleTextContent(ctx.dom.secondarySubRoot);
  if (secondaryHasContent) {
    const secondaryRect = toMeasuredRect(ctx.dom.secondarySubContainer.getBoundingClientRect());
    if (secondaryRect) {
      combinedRect = combinedRect ? unionRects(combinedRect, secondaryRect) : secondaryRect;
    }
  }

  if (!combinedRect) {
    return null;
  }

  return {
    x: combinedRect.x,
    y: combinedRect.y,
    width: round2(Math.max(0, combinedRect.width)),
    height: round2(Math.max(0, combinedRect.height)),
  };
}

export function createOverlayContentMeasurementReporter(ctx: RendererContext) {
  let debounceTimer: number | null = null;

  function emitNow(): void {
    if (!isMeasurableOverlayLayer(ctx.platform.overlayLayer)) {
      return;
    }

    const measurement: OverlayContentMeasurement = {
      layer: ctx.platform.overlayLayer,
      measuredAtMs: Date.now(),
      viewport: {
        width: window.innerWidth,
        height: window.innerHeight,
      },
      // Explicit null rect signals "no content yet", and main should use fallback bounds.
      contentRect: collectContentRect(ctx),
    };

    window.electronAPI.reportOverlayContentBounds(measurement);
  }

  function schedule(): void {
    if (debounceTimer !== null) {
      window.clearTimeout(debounceTimer);
    }
    debounceTimer = window.setTimeout(() => {
      debounceTimer = null;
      emitNow();
    }, MEASUREMENT_DEBOUNCE_MS);
  }

  return {
    emitNow,
    schedule,
  };
}
