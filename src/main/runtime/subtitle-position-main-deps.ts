import type {
  createLoadSubtitlePositionHandler,
  createSaveSubtitlePositionHandler,
} from './subtitle-position';

type LoadSubtitlePositionMainDeps = Parameters<typeof createLoadSubtitlePositionHandler>[0];
type SaveSubtitlePositionMainDeps = Parameters<typeof createSaveSubtitlePositionHandler>[0];

export function createBuildLoadSubtitlePositionMainDepsHandler(deps: LoadSubtitlePositionMainDeps) {
  return (): LoadSubtitlePositionMainDeps => ({
    loadSubtitlePositionCore: () => deps.loadSubtitlePositionCore(),
    setSubtitlePosition: (position) => deps.setSubtitlePosition(position),
  });
}

export function createBuildSaveSubtitlePositionMainDepsHandler(deps: SaveSubtitlePositionMainDeps) {
  return (): SaveSubtitlePositionMainDeps => ({
    saveSubtitlePositionCore: (position) => deps.saveSubtitlePositionCore(position),
    setSubtitlePosition: (position) => deps.setSubtitlePosition(position),
  });
}
