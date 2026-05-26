import { exportLogsArchiveForCurrentUser } from '../../src/main/runtime/log-export.js';
import type { ExportLogsResult } from '../../src/main/runtime/log-export.js';
import type { LauncherCommandContext } from './context.js';

interface LogsCommandDeps {
  exportLogsArchive(): ExportLogsResult;
}

const defaultDeps: LogsCommandDeps = {
  exportLogsArchive: () => exportLogsArchiveForCurrentUser(),
};

export function runLogsCommand(
  context: LauncherCommandContext,
  deps: LogsCommandDeps = defaultDeps,
): boolean {
  if (!context.args.logsExport) {
    return false;
  }

  const result = deps.exportLogsArchive();
  context.processAdapter.writeStdout(`${result.zipPath}\n`);
  return true;
}
