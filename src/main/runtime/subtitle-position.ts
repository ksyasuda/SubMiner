import type { SubtitlePosition } from '../../types';

export function createLoadSubtitlePositionHandler(deps: {
  loadSubtitlePositionCore: () => SubtitlePosition | null;
  setSubtitlePosition: (position: SubtitlePosition | null) => void;
}) {
  return (): SubtitlePosition | null => {
    const position = deps.loadSubtitlePositionCore();
    deps.setSubtitlePosition(position);
    return position;
  };
}

export function createSaveSubtitlePositionHandler(deps: {
  saveSubtitlePositionCore: (position: SubtitlePosition) => void;
  setSubtitlePosition: (position: SubtitlePosition) => void;
}) {
  return (position: SubtitlePosition): void => {
    deps.setSubtitlePosition(position);
    deps.saveSubtitlePositionCore(position);
  };
}
