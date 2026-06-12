import type { MessageBoxOptions, MessageBoxReturnValue } from 'electron';

type ShowMessageBox = (options: MessageBoxOptions) => Promise<MessageBoxReturnValue>;

export async function showLogExportSuccessDialog(options: {
  zipPath: string;
  showMessageBox: ShowMessageBox;
  showItemInFolder: (path: string) => void;
  logWarn: (message: string, details?: unknown) => void;
}): Promise<void> {
  const successDialog = await options
    .showMessageBox({
      type: 'info',
      title: 'SubMiner logs exported',
      message: 'SubMiner log export created.',
      detail: options.zipPath,
      buttons: ['OK', 'Show in Folder'],
      defaultId: 0,
      cancelId: 0,
    })
    .catch((dialogError) => {
      options.logWarn('Failed to show log export success dialog.', dialogError);
      return undefined;
    });

  if (successDialog?.response === 1) {
    options.showItemInFolder(options.zipPath);
  }
}

export async function showLogExportErrorDialog(options: {
  message: string;
  showMessageBox: ShowMessageBox;
  logWarn: (message: string, details?: unknown) => void;
}): Promise<void> {
  await options
    .showMessageBox({
      type: 'error',
      title: 'SubMiner log export failed',
      message: 'Could not export SubMiner logs.',
      detail: options.message,
    })
    .catch((dialogError) => {
      options.logWarn('Failed to show log export error dialog.', dialogError);
    });
}
