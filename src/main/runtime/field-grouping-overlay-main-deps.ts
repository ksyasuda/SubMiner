import type { FieldGroupingOverlayRuntimeOptions } from '../../core/services/field-grouping-overlay';

type FieldGroupingOverlayMainDeps<TModal extends string> = Omit<
  FieldGroupingOverlayRuntimeOptions<TModal>,
  'sendToVisibleOverlay'
> & {
  sendToActiveOverlayWindow: (
    channel: string,
    payload?: unknown,
    runtimeOptions?: { restoreOnModalClose?: TModal; preferModalWindow?: boolean },
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
    // These are optional on the runtime options, so a missing forward compiles silently — but
    // dropping them left the field grouping modal with no modal-open ack/retry, no teardown on
    // failure, and no warn logging, which is why the earlier recovery fixes never took effect.
    waitForModalOpen: deps.waitForModalOpen,
    handleOverlayModalClosed: deps.handleOverlayModalClosed,
    logWarn: deps.logWarn,
    ensureOverlayStartupPrereqs: deps.ensureOverlayStartupPrereqs,
    ensureOverlayWindowsReadyForVisibilityActions:
      deps.ensureOverlayWindowsReadyForVisibilityActions,
    sendToVisibleOverlay: (
      channel: string,
      payload?: unknown,
      runtimeOptions?: { restoreOnModalClose?: TModal; preferModalWindow?: boolean },
    ) => deps.sendToActiveOverlayWindow(channel, payload, runtimeOptions),
  });
}
