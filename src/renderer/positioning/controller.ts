import type { RendererContext } from '../context';
import {
  createInMemorySubtitlePositionController,
  type SubtitlePositionController,
} from './position-state.js';

export function createPositioningController(ctx: RendererContext): SubtitlePositionController {
  return createInMemorySubtitlePositionController(ctx);
}
