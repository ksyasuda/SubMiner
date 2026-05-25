export type UpdateAvailableChoice = 'update' | 'close';
export type RestartChoice = 'restart' | 'later';

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
    title: 'SubMiner Updates',
    message: `SubMiner is up to date (v${version})`,
    buttons: ['Close'],
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
    title: 'SubMiner Updates',
    message: `SubMiner v${version} is available`,
    buttons: ['Update', 'Close'],
    defaultId: 0,
    cancelId: 1,
  });
  return result.response === 0 ? 'update' : 'close';
}

export async function showRestartDialog(showMessageBox: ShowMessageBox): Promise<RestartChoice> {
  const result = await showMessageBox({
    type: 'question',
    title: 'SubMiner Updates',
    message: 'Restart to update',
    buttons: ['Restart', 'Later'],
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
    title: 'SubMiner Updates',
    message: 'Manual install required',
    detail: `SubMiner v${version} is available, but this build cannot install app updates automatically. Download and install the latest release, then reopen SubMiner.`,
    buttons: ['Close'],
  });
}

export async function showUpdateFailedDialog(
  showMessageBox: ShowMessageBox,
  message: string,
): Promise<void> {
  await showMessageBox({
    type: 'error',
    title: 'SubMiner Updates',
    message: 'Update check failed',
    detail: message,
    buttons: ['Close'],
  });
}
