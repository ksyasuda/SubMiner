import { app, dialog, shell } from 'electron';
import * as os from 'os';
import { showLogExportErrorDialog, showLogExportSuccessDialog } from './log-export-dialogs';
import { exportLogsArchive } from './log-export';

export interface LogExportTrayRuntimeDeps {
  flushMpvLog: () => Promise<void>;
  logInfo: (message: string) => void;
  logWarn: (message: string, details?: unknown) => void;
}

export function createLogExportTrayRuntime(deps: LogExportTrayRuntimeDeps): {
  exportLogsFromTray: () => Promise<void>;
} {
  function describeUnknownError(error: unknown): string {
    return error instanceof Error ? error.message : String(error);
  }

  async function exportLogsFromTray(): Promise<void> {
    try {
      await deps.flushMpvLog();
    } catch (error) {
      deps.logWarn('Failed to flush mpv log before exporting logs from tray.', error);
    }

    try {
      const result = exportLogsArchive({
        platform: process.platform,
        homeDir: os.homedir(),
        appDataDir: app.getPath('appData'),
      });
      deps.logInfo(
        `Exported ${result.exportedFiles.length} sanitized log file(s) to ${result.zipPath}`,
      );
      await showLogExportSuccessDialog({
        zipPath: result.zipPath,
        showMessageBox: (options) => dialog.showMessageBox(options),
        showItemInFolder: (zipPath) => shell.showItemInFolder(zipPath),
        logWarn: deps.logWarn,
      });
    } catch (error) {
      const message = describeUnknownError(error);
      deps.logWarn('Failed to export logs from tray.', error);
      await showLogExportErrorDialog({
        message,
        showMessageBox: (options) => dialog.showMessageBox(options),
        logWarn: deps.logWarn,
      });
    }
  }

  return { exportLogsFromTray };
}
