export type OverlayLayer = 'visible' | 'modal';

export type PlatformInfo = {
  overlayLayer: OverlayLayer;
  isModalLayer: boolean;
  isLinuxPlatform: boolean;
  isMacOSPlatform: boolean;
  shouldToggleMouseIgnore: boolean;
};

export function resolvePlatformInfo(): PlatformInfo {
  const overlayLayerFromPreload = window.electronAPI.getOverlayLayer();
  const queryLayer = new URLSearchParams(window.location.search).get('layer');
  const overlayLayerFromQuery: OverlayLayer | null =
    queryLayer === 'visible' || queryLayer === 'modal' ? queryLayer : null;

  const overlayLayer: OverlayLayer =
    overlayLayerFromQuery ??
    (overlayLayerFromPreload === 'visible' || overlayLayerFromPreload === 'modal'
      ? overlayLayerFromPreload
      : 'visible');

  const isModalLayer = overlayLayer === 'modal';
  const isLinuxPlatform = navigator.platform.toLowerCase().includes('linux');
  const isMacOSPlatform =
    navigator.platform.toLowerCase().includes('mac') || /mac/i.test(navigator.userAgent);

  return {
    overlayLayer,
    isModalLayer,
    isLinuxPlatform,
    isMacOSPlatform,
    shouldToggleMouseIgnore: !isLinuxPlatform && !isModalLayer,
  };
}
