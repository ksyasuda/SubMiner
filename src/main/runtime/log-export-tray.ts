import { app, dialog, shell } from 'electron';
import * as os from 'os';
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
      void dialog
        .showMessageBox({
          type: 'info',
          title: 'SubMiner logs exported',
          message: 'SubMiner log export created.',
          detail: result.zipPath,
          buttons: ['OK', 'Show in Folder'],
          defaultId: 0,
          cancelId: 0,
        })
        .then((response) => {
          if (response.response === 1) {
            shell.showItemInFolder(result.zipPath);
          }
        });
    } catch (error) {
      const message = describeUnknownError(error);
      deps.logWarn('Failed to export logs from tray.', error);
      void dialog.showMessageBox({
        type: 'error',
        title: 'SubMiner log export failed',
        message: 'Could not export SubMiner logs.',
        detail: message,
      });
    }
  }

  return { exportLogsFromTray };
}
