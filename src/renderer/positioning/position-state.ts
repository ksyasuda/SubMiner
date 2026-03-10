import type { SubtitlePosition } from '../../types';
import type { RendererContext } from '../context';

const PREFERRED_Y_PERCENT_MIN = 2;
const PREFERRED_Y_PERCENT_MAX = 80;
const SUBTITLE_EDGE_PADDING_PX = 12;

export type SubtitlePositionController = {
  applyStoredSubtitlePosition: (position: SubtitlePosition | null, source: string) => void;
  getCurrentYPercent: () => number;
  applyYPercent: (yPercent: number) => void;
  persistSubtitlePositionPatch: (patch: Partial<SubtitlePosition>) => void;
};

function getViewportHeight(): number {
  return Math.max(window.innerHeight || 0, 1);
}

function getSubtitleContainerHeight(ctx: RendererContext): number {
  const container = ctx.dom.subtitleContainer as HTMLElement & {
    offsetHeight?: number;
    getBoundingClientRect?: () => { height?: number };
  };
  if (typeof container.offsetHeight === 'number' && Number.isFinite(container.offsetHeight)) {
    return Math.max(container.offsetHeight, 0);
  }
  if (typeof container.getBoundingClientRect === 'function') {
    const height = container.getBoundingClientRect().height;
    if (typeof height === 'number' && Number.isFinite(height)) {
      return Math.max(height, 0);
    }
  }
  return 0;
}

function resolveYPercentClampRange(ctx: RendererContext): { min: number; max: number } {
  const viewportHeight = getViewportHeight();
  const subtitleHeight = getSubtitleContainerHeight(ctx);
  const minPercent = Math.max(
    PREFERRED_Y_PERCENT_MIN,
    (SUBTITLE_EDGE_PADDING_PX / viewportHeight) * 100,
  );
  const maxMarginBottomPx = Math.max(
    SUBTITLE_EDGE_PADDING_PX,
    viewportHeight - subtitleHeight - SUBTITLE_EDGE_PADDING_PX,
  );
  const maxPercent = Math.min(PREFERRED_Y_PERCENT_MAX, (maxMarginBottomPx / viewportHeight) * 100);

  if (maxPercent < minPercent) {
    return { min: minPercent, max: minPercent };
  }

  return { min: minPercent, max: maxPercent };
}

function clampYPercent(ctx: RendererContext, yPercent: number): number {
  const { min, max } = resolveYPercentClampRange(ctx);
  return Math.max(min, Math.min(max, yPercent));
}

function getPersistedYPercent(ctx: RendererContext, position: SubtitlePosition | null): number {
  if (!position || typeof position.yPercent !== 'number' || !Number.isFinite(position.yPercent)) {
    return ctx.state.persistedSubtitlePosition.yPercent;
  }

  return position.yPercent;
}

function updatePersistedSubtitlePosition(
  ctx: RendererContext,
  position: SubtitlePosition | null,
): void {
  ctx.state.persistedSubtitlePosition = {
    yPercent: getPersistedYPercent(ctx, position),
  };
}

function getNextPersistedPosition(
  ctx: RendererContext,
  patch: Partial<SubtitlePosition>,
): SubtitlePosition {
  return {
    yPercent:
      typeof patch.yPercent === 'number' && Number.isFinite(patch.yPercent)
        ? patch.yPercent
        : ctx.state.persistedSubtitlePosition.yPercent,
  };
}

function applyMarginBottom(ctx: RendererContext, yPercent: number): void {
  const clampedPercent = clampYPercent(ctx, yPercent);
  ctx.state.currentYPercent = clampedPercent;
  const marginBottom = (clampedPercent / 100) * getViewportHeight();

  ctx.dom.subtitleContainer.style.position = '';
  ctx.dom.subtitleContainer.style.left = '';
  ctx.dom.subtitleContainer.style.top = '';
  ctx.dom.subtitleContainer.style.right = '';
  ctx.dom.subtitleContainer.style.transform = '';
  ctx.dom.subtitleContainer.style.marginBottom = `${marginBottom}px`;
}

export function createInMemorySubtitlePositionController(
  ctx: RendererContext,
): SubtitlePositionController {
  function getCurrentYPercent(): number {
    if (ctx.state.currentYPercent !== null) {
      return ctx.state.currentYPercent;
    }

    const marginBottom = parseFloat(ctx.dom.subtitleContainer.style.marginBottom) || 60;
    ctx.state.currentYPercent = clampYPercent(ctx, (marginBottom / getViewportHeight()) * 100);
    return ctx.state.currentYPercent;
  }

  function applyYPercent(yPercent: number): void {
    applyMarginBottom(ctx, yPercent);
  }

  function persistSubtitlePositionPatch(patch: Partial<SubtitlePosition>): void {
    const nextPosition = getNextPersistedPosition(ctx, patch);
    ctx.state.persistedSubtitlePosition = nextPosition;
    window.electronAPI.saveSubtitlePosition(nextPosition);
  }

  function applyStoredSubtitlePosition(position: SubtitlePosition | null, source: string): void {
    updatePersistedSubtitlePosition(ctx, position);
    if (position && position.yPercent !== undefined) {
      applyYPercent(position.yPercent);
      console.log('Applied subtitle position from', source, ':', position.yPercent, '%');
      return;
    }

    const defaultMarginBottom = 60;
    const defaultYPercent = (defaultMarginBottom / getViewportHeight()) * 100;
    applyYPercent(defaultYPercent);
    console.log('Applied default subtitle position from', source);
  }

  return {
    applyStoredSubtitlePosition,
    getCurrentYPercent,
    applyYPercent,
    persistSubtitlePositionPatch,
  };
}
