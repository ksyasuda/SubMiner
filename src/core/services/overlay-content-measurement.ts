import { OverlayContentMeasurement, OverlayContentRect, OverlayLayer } from '../../types';
import { createLogger } from '../../logger';

const logger = createLogger('main:overlay-content-measurement');
const MAX_VIEWPORT = 10000;
const MAX_RECT_DIMENSION = 10000;
const MAX_RECT_OFFSET = 50000;
const MAX_FUTURE_TIMESTAMP_MS = 60_000;
const INVALID_LOG_THROTTLE_MS = 10_000;

type OverlayMeasurementStore = Record<OverlayLayer, OverlayContentMeasurement | null>;

export function sanitizeOverlayContentMeasurement(
  payload: unknown,
  nowMs: number,
): OverlayContentMeasurement | null {
  if (!payload || typeof payload !== 'object') return null;

  const candidate = payload as {
    layer?: unknown;
    measuredAtMs?: unknown;
    viewport?: { width?: unknown; height?: unknown };
    contentRect?: {
      x?: unknown;
      y?: unknown;
      width?: unknown;
      height?: unknown;
    } | null;
  };

  if (candidate.layer !== 'visible') {
    return null;
  }

  const viewportWidth = readFiniteInRange(candidate.viewport?.width, 1, MAX_VIEWPORT);
  const viewportHeight = readFiniteInRange(candidate.viewport?.height, 1, MAX_VIEWPORT);

  if (!Number.isFinite(viewportWidth) || !Number.isFinite(viewportHeight)) {
    return null;
  }

  const measuredAtMs = readFiniteInRange(
    candidate.measuredAtMs,
    1,
    nowMs + MAX_FUTURE_TIMESTAMP_MS,
  );
  if (!Number.isFinite(measuredAtMs)) {
    return null;
  }

  const contentRect = sanitizeOverlayContentRect(candidate.contentRect);
  if (candidate.contentRect !== null && !contentRect) {
    return null;
  }

  return {
    layer: candidate.layer,
    measuredAtMs,
    viewport: { width: viewportWidth, height: viewportHeight },
    contentRect,
  };
}

function sanitizeOverlayContentRect(rect: unknown): OverlayContentRect | null {
  if (rect === null || rect === undefined) {
    return null;
  }

  if (!rect || typeof rect !== 'object') {
    return null;
  }

  const candidate = rect as {
    x?: unknown;
    y?: unknown;
    width?: unknown;
    height?: unknown;
  };

  const width = readFiniteInRange(candidate.width, 0, MAX_RECT_DIMENSION);
  const height = readFiniteInRange(candidate.height, 0, MAX_RECT_DIMENSION);
  const x = readFiniteInRange(candidate.x, -MAX_RECT_OFFSET, MAX_RECT_OFFSET);
  const y = readFiniteInRange(candidate.y, -MAX_RECT_OFFSET, MAX_RECT_OFFSET);

  if (
    !Number.isFinite(width) ||
    !Number.isFinite(height) ||
    !Number.isFinite(x) ||
    !Number.isFinite(y)
  ) {
    return null;
  }

  return { x, y, width, height };
}

function readFiniteInRange(value: unknown, min: number, max: number): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    return Number.NaN;
  }
  if (value < min || value > max) {
    return Number.NaN;
  }
  return value;
}

export function createOverlayContentMeasurementStore(options?: {
  now?: () => number;
  warn?: (message: string) => void;
}) {
  const now = options?.now ?? (() => Date.now());
  const warn = options?.warn ?? ((message: string) => logger.warn(message));
  const latestByLayer: OverlayMeasurementStore = {
    visible: null,
  };

  let droppedInvalid = 0;
  let lastInvalidLogAtMs = 0;

  function report(payload: unknown): OverlayContentMeasurement | null {
    const nowMs = now();
    const measurement = sanitizeOverlayContentMeasurement(payload, nowMs);
    if (!measurement) {
      droppedInvalid += 1;
      if (droppedInvalid > 0 && nowMs - lastInvalidLogAtMs >= INVALID_LOG_THROTTLE_MS) {
        warn(
          `[overlay-content-bounds] Dropped ${droppedInvalid} invalid measurement payload(s) in the last ${INVALID_LOG_THROTTLE_MS}ms.`,
        );
        droppedInvalid = 0;
        lastInvalidLogAtMs = nowMs;
      }
      return null;
    }

    latestByLayer[measurement.layer] = measurement;
    return measurement;
  }

  function getLatestByLayer(layer: OverlayLayer): OverlayContentMeasurement | null {
    return latestByLayer[layer];
  }

  return {
    getLatestByLayer,
    report,
  };
}
