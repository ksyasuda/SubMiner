import type { MessageBoxOptions, MessageBoxReturnValue } from 'electron';
import { i18n } from '../../i18n/index.js';

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
      title: i18n.t('dialog.exportSuccess'),
      message: i18n.t('dialog.exportSuccessMsg'),
      detail: options.zipPath,
      buttons: [i18n.t('dialog.ok'), i18n.t('dialog.showInFolder')],
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
      title: i18n.t('dialog.exportFailed'),
      message: i18n.t('dialog.exportFailedMsg'),
      detail: options.message,
    })
    .catch((dialogError) => {
      options.logWarn('Failed to show log export error dialog.', dialogError);
    });
}
