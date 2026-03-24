export function resolveOverlayLayerFromArgv(
  argv: readonly string[] | null | undefined,
): 'visible' | 'modal' | null {
  const overlayLayerArg = argv?.find((arg) => arg.startsWith('--overlay-layer='));
  const overlayLayerFromArg = overlayLayerArg?.slice('--overlay-layer='.length);

  return overlayLayerFromArg === 'visible' || overlayLayerFromArg === 'modal'
    ? overlayLayerFromArg
    : null;
}
