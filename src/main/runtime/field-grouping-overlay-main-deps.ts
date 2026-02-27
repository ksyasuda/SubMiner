import type { FieldGroupingOverlayRuntimeOptions } from '../../core/services/field-grouping-overlay';

type FieldGroupingOverlayMainDeps<TModal extends string> = Omit<
  FieldGroupingOverlayRuntimeOptions<TModal>,
  'sendToVisibleOverlay'
> & {
  sendToActiveOverlayWindow: (
    channel: string,
    payload?: unknown,
    runtimeOptions?: { restoreOnModalClose?: TModal },
  ) => boolean;
};

type BuiltFieldGroupingOverlayMainDeps<TModal extends string> =
  FieldGroupingOverlayRuntimeOptions<TModal> & {
    sendToVisibleOverlay: NonNullable<
      FieldGroupingOverlayRuntimeOptions<TModal>['sendToVisibleOverlay']
    >;
  };

export function createBuildFieldGroupingOverlayMainDepsHandler<TModal extends string>(
  deps: FieldGroupingOverlayMainDeps<TModal>,
) {
  return (): BuiltFieldGroupingOverlayMainDeps<TModal> => ({
    getMainWindow: () => deps.getMainWindow(),
    getVisibleOverlayVisible: () => deps.getVisibleOverlayVisible(),
    setVisibleOverlayVisible: (visible: boolean) => deps.setVisibleOverlayVisible(visible),
    getResolver: () => deps.getResolver(),
    setResolver: (resolver) => deps.setResolver(resolver),
    getRestoreVisibleOverlayOnModalClose: () => deps.getRestoreVisibleOverlayOnModalClose(),
    sendToVisibleOverlay: (
      channel: string,
      payload?: unknown,
      runtimeOptions?: { restoreOnModalClose?: TModal },
    ) => deps.sendToActiveOverlayWindow(channel, payload, runtimeOptions),
  });
}
