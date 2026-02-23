import type { SubtitlePosition } from '../../types';
import type { RendererContext } from '../context';

const PREFERRED_Y_PERCENT_MIN = 2;
const PREFERRED_Y_PERCENT_MAX = 80;

export type SubtitlePositionController = {
  applyStoredSubtitlePosition: (position: SubtitlePosition | null, source: string) => void;
  getCurrentYPercent: () => number;
  applyYPercent: (yPercent: number) => void;
  persistSubtitlePositionPatch: (patch: Partial<SubtitlePosition>) => void;
};

function clampYPercent(yPercent: number): number {
  return Math.max(PREFERRED_Y_PERCENT_MIN, Math.min(PREFERRED_Y_PERCENT_MAX, yPercent));
}

function getPersistedYPercent(ctx: RendererContext, position: SubtitlePosition | null): number {
  if (!position || typeof position.yPercent !== 'number' || !Number.isFinite(position.yPercent)) {
    return ctx.state.persistedSubtitlePosition.yPercent;
  }

  return position.yPercent;
}

function getPersistedOffset(
  position: SubtitlePosition | null,
  key: 'invisibleOffsetXPx' | 'invisibleOffsetYPx',
): number {
  if (position && typeof position[key] === 'number' && Number.isFinite(position[key])) {
    return position[key];
  }

  return 0;
}

function updatePersistedSubtitlePosition(
  ctx: RendererContext,
  position: SubtitlePosition | null,
): void {
  ctx.state.persistedSubtitlePosition = {
    yPercent: getPersistedYPercent(ctx, position),
    invisibleOffsetXPx: getPersistedOffset(position, 'invisibleOffsetXPx'),
    invisibleOffsetYPx: getPersistedOffset(position, 'invisibleOffsetYPx'),
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
    invisibleOffsetXPx:
      typeof patch.invisibleOffsetXPx === 'number' && Number.isFinite(patch.invisibleOffsetXPx)
        ? patch.invisibleOffsetXPx
        : (ctx.state.persistedSubtitlePosition.invisibleOffsetXPx ?? 0),
    invisibleOffsetYPx:
      typeof patch.invisibleOffsetYPx === 'number' && Number.isFinite(patch.invisibleOffsetYPx)
        ? patch.invisibleOffsetYPx
        : (ctx.state.persistedSubtitlePosition.invisibleOffsetYPx ?? 0),
  };
}

export function createInMemorySubtitlePositionController(
  ctx: RendererContext,
): SubtitlePositionController {
  function getCurrentYPercent(): number {
    if (ctx.state.currentYPercent !== null) {
      return ctx.state.currentYPercent;
    }

    const marginBottom = parseFloat(ctx.dom.subtitleContainer.style.marginBottom) || 60;
    ctx.state.currentYPercent = clampYPercent((marginBottom / window.innerHeight) * 100);
    return ctx.state.currentYPercent;
  }

  function applyYPercent(yPercent: number): void {
    const clampedPercent = clampYPercent(yPercent);
    ctx.state.currentYPercent = clampedPercent;
    const marginBottom = (clampedPercent / 100) * window.innerHeight;

    ctx.dom.subtitleContainer.style.position = '';
    ctx.dom.subtitleContainer.style.left = '';
    ctx.dom.subtitleContainer.style.top = '';
    ctx.dom.subtitleContainer.style.right = '';
    ctx.dom.subtitleContainer.style.transform = '';
    ctx.dom.subtitleContainer.style.marginBottom = `${marginBottom}px`;
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
    const defaultYPercent = (defaultMarginBottom / window.innerHeight) * 100;
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
