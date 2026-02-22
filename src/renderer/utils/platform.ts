export type OverlayLayer = 'visible' | 'invisible';

export type PlatformInfo = {
  overlayLayer: OverlayLayer;
  isInvisibleLayer: boolean;
  isLinuxPlatform: boolean;
  isMacOSPlatform: boolean;
  shouldToggleMouseIgnore: boolean;
  invisiblePositionEditToggleCode: string;
  invisiblePositionStepPx: number;
  invisiblePositionStepFastPx: number;
};

export function resolvePlatformInfo(): PlatformInfo {
  const overlayLayerFromPreload = window.electronAPI.getOverlayLayer();
  const queryLayer = new URLSearchParams(window.location.search).get('layer');
  const overlayLayerFromQuery: OverlayLayer | null =
    queryLayer === 'visible' || queryLayer === 'invisible' ? queryLayer : null;

  const overlayLayer: OverlayLayer =
    overlayLayerFromQuery ??
    (overlayLayerFromPreload === 'visible' || overlayLayerFromPreload === 'invisible'
      ? overlayLayerFromPreload
      : 'visible');

  const isInvisibleLayer = overlayLayer === 'invisible';
  const isLinuxPlatform = navigator.platform.toLowerCase().includes('linux');
  const isMacOSPlatform =
    navigator.platform.toLowerCase().includes('mac') || /mac/i.test(navigator.userAgent);

  return {
    overlayLayer,
    isInvisibleLayer,
    isLinuxPlatform,
    isMacOSPlatform,
    shouldToggleMouseIgnore: !isLinuxPlatform,
    invisiblePositionEditToggleCode: 'KeyP',
    invisiblePositionStepPx: 1,
    invisiblePositionStepFastPx: 4,
  };
}
