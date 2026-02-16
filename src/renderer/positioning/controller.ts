import type { ModalStateReader, RendererContext } from "../context";
import {
  createInMemorySubtitlePositionController,
  type SubtitlePositionController,
} from "./position-state.js";
import {
  createInvisibleOffsetController,
  type InvisibleOffsetController,
} from "./invisible-offset.js";
import {
  createMpvSubtitleLayoutController,
  type MpvSubtitleLayoutController,
} from "./invisible-layout.js";

type PositioningControllerOptions = {
  modalStateReader: Pick<ModalStateReader, "isAnySettingsModalOpen">;
  applySubtitleFontSize: (fontSize: number) => void;
};

export function createPositioningController(
  ctx: RendererContext,
  options: PositioningControllerOptions,
) {
  const visible = createInMemorySubtitlePositionController(ctx);
  const invisibleOffset = createInvisibleOffsetController(
    ctx,
    options.modalStateReader,
  );
  const invisibleLayout = createMpvSubtitleLayoutController(
    ctx,
    options.applySubtitleFontSize,
    {
      applyInvisibleSubtitleOffsetPosition:
        invisibleOffset.applyInvisibleSubtitleOffsetPosition,
      updateInvisiblePositionEditHud: invisibleOffset.updateInvisiblePositionEditHud,
    },
  );

  return {
    ...visible,
    ...invisibleOffset,
    ...invisibleLayout,
  } as SubtitlePositionController &
    InvisibleOffsetController &
    MpvSubtitleLayoutController;
}
