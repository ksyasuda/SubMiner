export type UpdateAvailableChoice = 'update' | 'close';
export type RestartChoice = 'restart' | 'later';

import { i18n } from '../../../i18n/index.js';

export interface MessageBoxResultLike {
  response: number;
}

export type ShowMessageBox = (options: {
  type?: 'info' | 'warning' | 'error' | 'question';
  title?: string;
  message: string;
  detail?: string;
  buttons?: string[];
  defaultId?: number;
  cancelId?: number;
}) => Promise<MessageBoxResultLike>;

export interface UpdateDialogPresenterDeps {
  showMessageBox: ShowMessageBox;
  focusApp?: () => void | Promise<void>;
  yieldToRunLoop?: () => Promise<void>;
  withStatsWindowLayerSuspended?: <T>(showDialog: () => Promise<T>) => Promise<T>;
  platform?: NodeJS.Platform;
}

export async function showNoUpdateDialog(
  showMessageBox: ShowMessageBox,
  version: string,
): Promise<void> {
  await showMessageBox({
    type: 'info',
    title: i18n.t('update.title'),
    message: i18n.t('update.upToDate', { version }),
    buttons: [i18n.t('dialog.close')],
  });
}

async function maybeFocusAppForDialog(deps: UpdateDialogPresenterDeps): Promise<void> {
  if ((deps.platform ?? process.platform) !== 'darwin') return;
  await deps.focusApp?.();
  // Yield to the macOS run loop so the activation request is processed before the
  // modal alert blocks JS execution; without this, the alert often appears behind
  // other apps when SubMiner is not the active app at dialog-show time.
  const yieldToRunLoop = deps.yieldToRunLoop ?? (() => new Promise((r) => setTimeout(r, 0)));
  await yieldToRunLoop();
}

export function createUpdateDialogPresenter(deps: UpdateDialogPresenterDeps) {
  const showFocusedMessageBox: ShowMessageBox = async (options) => {
    const showDialog = async (): Promise<MessageBoxResultLike> => {
      try {
        await maybeFocusAppForDialog(deps);
      } catch {
        // Best-effort focus only; never block the dialog itself.
      }
      return deps.showMessageBox(options);
    };

    return deps.withStatsWindowLayerSuspended
      ? deps.withStatsWindowLayerSuspended(showDialog)
      : showDialog();
  };

  return {
    showNoUpdateDialog: (version: string) => showNoUpdateDialog(showFocusedMessageBox, version),
    showUpdateAvailableDialog: (version: string) =>
      showUpdateAvailableDialog(showFocusedMessageBox, version),
    showUpdateFailedDialog: (message: string) =>
      showUpdateFailedDialog(showFocusedMessageBox, message),
    showManualUpdateRequiredDialog: (version: string) =>
      showManualUpdateRequiredDialog(showFocusedMessageBox, version),
    showRestartDialog: () => showRestartDialog(showFocusedMessageBox),
  };
}

export async function showUpdateAvailableDialog(
  showMessageBox: ShowMessageBox,
  version: string,
): Promise<UpdateAvailableChoice> {
  const result = await showMessageBox({
    type: 'question',
    title: i18n.t('update.title'),
    message: i18n.t('update.available', { version }),
    buttons: [i18n.t('update.updateButton'), i18n.t('dialog.close')],
    defaultId: 0,
    cancelId: 1,
  });
  return result.response === 0 ? 'update' : 'close';
}

export async function showRestartDialog(showMessageBox: ShowMessageBox): Promise<RestartChoice> {
  const result = await showMessageBox({
    type: 'question',
    title: i18n.t('update.title'),
    message: i18n.t('update.restart'),
    buttons: [i18n.t('update.restartButton'), i18n.t('update.laterButton')],
    defaultId: 0,
    cancelId: 1,
  });
  return result.response === 0 ? 'restart' : 'later';
}

export async function showManualUpdateRequiredDialog(
  showMessageBox: ShowMessageBox,
  version: string,
): Promise<void> {
  await showMessageBox({
    type: 'warning',
    title: i18n.t('update.title'),
    message: i18n.t('update.manualRequired'),
    detail: i18n.t('update.manualDetail', { version }),
    buttons: [i18n.t('dialog.close')],
  });
}

export async function showUpdateFailedDialog(
  showMessageBox: ShowMessageBox,
  message: string,
): Promise<void> {
  await showMessageBox({
    type: 'error',
    title: i18n.t('update.title'),
    message: i18n.t('update.failed'),
    detail: message,
    buttons: [i18n.t('dialog.close')],
  });
}
