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
  focusApp?: () => void;
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

function maybeFocusAppForDialog(deps: UpdateDialogPresenterDeps): void {
  if ((deps.platform ?? process.platform) !== 'darwin') return;
  deps.focusApp?.();
}

export function createUpdateDialogPresenter(deps: UpdateDialogPresenterDeps) {
  const showFocusedMessageBox: ShowMessageBox = async (options) => {
    maybeFocusAppForDialog(deps);
    return deps.showMessageBox(options);
  };

  return {
    showNoUpdateDialog: (version: string) => showNoUpdateDialog(showFocusedMessageBox, version),
    showUpdateAvailableDialog: (version: string) =>
      showUpdateAvailableDialog(showFocusedMessageBox, version),
    showUpdateFailedDialog: (message: string) =>
      showUpdateFailedDialog(showFocusedMessageBox, message),
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
